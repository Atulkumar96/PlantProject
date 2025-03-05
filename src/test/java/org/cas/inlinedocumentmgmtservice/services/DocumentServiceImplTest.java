package org.cas.inlinedocumentmgmtservice.services;

// Static imports for JUnit assertions and Mockito methods,
// so we can use methods like assertEquals, when, etc. without prefixing them.
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;                     // DTO for plant-related data.
import org.cas.inlinedocumentmgmtservice.exceptions.DocumentProcessingException; // Custom exception for document errors.
import org.junit.jupiter.api.BeforeEach;                                     // JUnit annotation to run a method before each test.
import org.junit.jupiter.api.Test;                                           // JUnit annotation for test methods.
import org.mockito.InjectMocks;                                              // Annotation to inject mocks into our test target.
import org.mockito.Mock;                                                     // Annotation to declare mock objects.
import org.mockito.MockitoAnnotations;                                       // Utility to initialize mocks.
import org.springframework.core.io.ResourceLoader;                          // Spring resource loader, likely used by DocumentServiceImpl.
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;                      // Represents an uploaded file.

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;                                                 // Used for input stream operations.
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class DocumentServiceImplTest {
    // Declare a mock for ResourceLoader, which may be used in DocumentServiceImpl.
    @Mock
    private ResourceLoader resourceLoader;

    // Declare a mock for MultipartFile to simulate file uploads in tests.
    @Mock
    private MultipartFile mockFile;

    // Inject the above mocks into the DocumentServiceImpl instance.
    @InjectMocks
    private DocumentServiceImpl documentServiceImpl;

    // Setup method that runs before each test, initializing all @Mock annotated fields.
    @BeforeEach
    void setUp() {
        MockitoAnnotations.openMocks(this);
    }

    /**
     * Test the importDocument() method.
     * This method should return a non-null WordDocument object.
     */
    @Test
    void testImportDocument() {
        // Act: Call the method to import the document from resources.
        Object wordDoc = documentServiceImpl.importDocument();
        // Assert: The returned document should not be null.
        assertNotNull(wordDoc, "The imported document should not be null.");
    }

    @Test
    void testMailMerge_Success() {
        // Create a PlantDto instance and set its plant name to "TestPlant".
        PlantDto plantDto = new PlantDto();
        plantDto.setPlant("TestPlant");

        // Verify that calling mailMerge on documentServiceImpl does not throw any exceptions.
        // This ensures that the method works as expected with valid input.
        assertDoesNotThrow(() -> documentServiceImpl.mailMerge(plantDto));
    }

    /**
     * Test the addComments(String, int) method.
     * We create a temporary minimal DOCX file, add comments, and ensure no exception is thrown.
     */
    @Test
    void testAddComments() throws Exception {
        // Create a temporary file representing a minimal DOCX.
        Path tempDoc = Files.createTempFile("testDocument", ".docx");
        Files.write(tempDoc, createMinimalValidDocx());

        // Act: Attempt to add 3 comments to the document.
        // (Note: verifying the actual comment content may require re-opening the document;
        // here we simply check that the method completes without exception.)
        documentServiceImpl.addComments(tempDoc.toString(), 3);

        // Cleanup: Delete the temporary file.
        Files.deleteIfExists(tempDoc);
    }

    @Test
    void testExtractReviewComments_WithFile_Success() throws IOException {
        // Arrange: Generate a minimal valid DOCX file using the helper method.
        byte[] validDocxBytes = createMinimalValidDocx();

        // Stub the behavior of the mock file's getInputStream() method -> to return our valid DOCX bytes.
        when(mockFile.getInputStream()).thenReturn(new ByteArrayInputStream(validDocxBytes));

        // Act: Call the method under test.
        String result = documentServiceImpl.extractReviewComments(mockFile);

        // Assert: Verify that the extraction result is not null.
        assertNotNull(result);
    }

    @Test
    void testExtractReviewComments_Failure() {
        // Stub the behavior of the mock file's isEmpty() method to simulate an empty file.
        when(mockFile.isEmpty()).thenReturn(true);

        // Verify that calling extractReviewComments with an empty file throws a DocumentProcessingException.
        // Since the file.getInputStream() will return null, the method should throw a null pointer exception.
        // We are not stubbing the behavior of getInputStream() here

        // Below assertion will confirm that the method properly handles invalid (empty) file input.
        assertThrows(DocumentProcessingException.class, () -> documentServiceImpl.extractReviewComments(mockFile));
    }

    /**
     * Test protectDocument(MultipartFile) method.
     * We use a mock MultipartFile containing minimal valid DOCX bytes.
     */
    @Test
    void testProtectDocument_WithMultipartFile() throws Exception {
        // Arrange: Prepare minimal valid DOCX bytes.
        byte[] docxBytes = createMinimalValidDocx();
        MultipartFile multipartFile = new MockMultipartFile("file", "test.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", docxBytes);

        // Act & Assert: The method should complete without throwing an exception.
        assertDoesNotThrow(() -> documentServiceImpl.protectDocument(multipartFile));
    }

    /**
     * Test insertImage(String, String, String) method.
     * We create a dummy DOCX and a dummy image (using minimal PNG header bytes), then check for an output file.
     */
    @Test
    void testInsertImage_WithPaths() throws Exception {
        // Create temporary input DOCX.
        Path inputDoc = Files.createTempFile("testDocument", ".docx");
        Files.write(inputDoc, createMinimalValidDocxWithTargetText());

        // Create a temporary dummy image file with a minimal PNG header.
        Path imageFile = Files.createTempFile("dummyImage", ".png");
        byte[] pngHeader = new byte[]{(byte)0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};
        Files.write(imageFile, pngHeader);

        // Define a temporary output file path.
        Path outputDoc = Files.createTempFile("outputDocument", ".docx");
        Files.deleteIfExists(outputDoc);

        // Act: Insert the image into the document.
        documentServiceImpl.insertImage(inputDoc.toString(), imageFile.toString(), outputDoc.toString());

        // Assert: Check that the output file now exists.
        assertTrue(Files.exists(outputDoc), "Output document should be created after image insertion.");

        // Cleanup temporary files.
        Files.deleteIfExists(inputDoc);
        Files.deleteIfExists(imageFile);
        Files.deleteIfExists(outputDoc);
    }

    /**
     * Test insertImage(MultipartFile, MultipartFile) method.
     * We simulate both document and image uploads using MockMultipartFile.
     */
    @Test
    void testInsertImage_WithMultipartFiles() throws Exception {
        // Arrange: Prepare minimal valid DOCX and dummy image bytes.
        byte[] docxBytes = createMinimalValidDocxWithTargetText();
        byte[] imageBytes = new byte[]{(byte)0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};

        MultipartFile docFile = new MockMultipartFile("docFile", "test.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", docxBytes);
        MultipartFile imageFile = new MockMultipartFile("imageFile", "dummy.png", "image/png", imageBytes);

        // Act: Call the insertImage method.
        String sfdt = documentServiceImpl.insertImage(docFile, imageFile);

        // Assert: Verify that an SFDT string is returned.
        assertNotNull(sfdt, "The returned SFDT string should not be null.");
        assertFalse(sfdt.isEmpty(), "The returned SFDT string should not be empty.");
    }


    @Test
    void testAppendSignature_Success() throws IOException {
        // Arrange: Generate a minimal valid DOCX file using the helper method.
        byte[] validDocxBytes = createMinimalValidDocx();

        // Stub the behavior of the mock file to simulate a valid DOCX file.
        when(mockFile.getInputStream()).thenReturn(new ByteArrayInputStream(validDocxBytes));

        // Act: Call the appendSignature method with the valid DOCX and a sample signature.
        String result = documentServiceImpl.appendSignature(mockFile, "Atul");

        // Assert: Check that the method returns a non-null result,
        // indicating that the file was processed and the signature appended.
        assertNotNull(result);
    }

    /**
     * Creates a minimal valid DOCX file in memory and returns its byte array.
     *
     * @return byte array representing a minimal DOCX file.
     * @throws IOException if an I/O error occurs.
     */
    private byte[] createMinimalValidDocx() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(baos)) {
            // Add [Content_Types].xml entry with minimal valid content.
            zos.putNextEntry(new ZipEntry("[Content_Types].xml"));
            String contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                    + "</Types>";
            zos.write(contentTypes.getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();

            // Add word/document.xml entry with minimal valid content.
            zos.putNextEntry(new ZipEntry("word/document.xml"));
            String documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                    + "<w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>";
            zos.write(documentXml.getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();
        }
        return baos.toByteArray();
    }

    /**
     * Helper method that creates a minimal valid DOCX file as a byte array.
     * The document XML includes the text "Subject Matter Experts", which is required
     * for the insertImage method to trigger image insertion.
     *
     * @return a byte array representing a minimal valid DOCX file.
     * @throws IOException if an I/O error occurs.
     */
    private byte[] createMinimalValidDocxWithTargetText() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(baos)) {
            // Create the [Content_Types].xml entry with minimal content.
            zos.putNextEntry(new ZipEntry("[Content_Types].xml"));
            String contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                    + "</Types>";
            zos.write(contentTypes.getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();

            // Create the word/document.xml entry with the required target text.
            zos.putNextEntry(new ZipEntry("word/document.xml"));
            String documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                    + "<w:body>"
                    + "<w:p><w:r><w:t>Subject Matter Experts</w:t></w:r></w:p>"
                    + "</w:body>"
                    + "</w:document>";
            zos.write(documentXml.getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();
        }
        return baos.toByteArray();
    }

}
