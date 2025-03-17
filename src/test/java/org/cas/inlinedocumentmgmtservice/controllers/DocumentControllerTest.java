package org.cas.inlinedocumentmgmtservice.controllers;

import static org.mockito.Mockito.*; // Mocking framework- mocking behaviors (when, doNothing, mock) to simulate the behavior of dependencies
import static org.junit.jupiter.api.Assertions.*; // to validate expected behavior in tests

// import DTOs
import com.syncfusion.docio.FormatType;
import com.syncfusion.docio.WordDocument;
import com.syncfusion.ej2.wordprocessor.WordProcessorHelper;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.dtos.ResponseDto;

// import services
import org.cas.inlinedocumentmgmtservice.dtos.SaveDto;
import org.cas.inlinedocumentmgmtservice.services.DocumentService;
import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;

// JUnit 5 (BeforeEach, Test): Defines test setup and test methods
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

// Mockito annotations (@Mock, @InjectMocks): Used to mock dependencies and inject them into the class under test
import org.mockito.*;

// Spring classes (ResponseEntity, HttpStatus): Used for HTTP responses in controller tests
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;

// MultipartFile: Represents an uploaded file in a request
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Unit test class for DocumentController
 */
public class DocumentControllerTest {
    /**
     * @Mock annotation creates mock objects for DocumentService and DocumentServiceImpl,
     * so that the actual implementations are not used in the tests.
     */
    @Mock
    private DocumentService documentService;

    @Mock
    private DocumentServiceImpl documentServiceImpl;

    /**
     * @InjectMocks annotation injects the mocked dependencies (documentService, documentServiceImpl)
     * into the DocumentController.
     */
    @InjectMocks
    private DocumentController documentController;

    /**
     * This method is executed before each test.
     * MockitoAnnotations.openMocks(this) initializes the annotated mocks (@Mock) before each test
     */
    @BeforeEach
    void setUp() {
        MockitoAnnotations.openMocks(this);
    }

    @Test
    void testMailMerge_Success() {
        // Arrange: Creates a sample DTO (PlantDto) that represents input data
        //PlantDto plantDto = new PlantDto();
        //plantDto.setPlant("TestPlant");

        String plantName = "TestPlant";

        // Mock: Mocks the behavior of the mailMerge() method in DocumentService
        doNothing().when(documentService).mailMerge(plantName);

        // Act: Calls the mailMerge method in DocumentController
        ResponseEntity<ResponseDto> response = documentController.mailMerge(plantName);

        // Assert: expected and actual values are compared
        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertEquals("Mail merge successful", response.getBody().getMessage());
    }

    @Test
    void testImportFile_Success() throws Exception {
        // Arrange: Creates a sample MultipartFile object
        MultipartFile mockFile = mock(MultipartFile.class);

        // Create a minimal valid DOCX file in memory using a ZIP stream.
        // A minimal DOCX must include at least a "[Content_Types].xml" and a "word/document.xml" entry.
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(baos)) {
            // Add the [Content_Types].xml entry with minimal valid content
            zos.putNextEntry(new ZipEntry("[Content_Types].xml"));
            String contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                    + "</Types>";
            zos.write(contentTypes.getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();

            // Add the word/document.xml entry with minimal valid content
            zos.putNextEntry(new ZipEntry("word/document.xml"));
            String documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                    + "<w:body></w:body></w:document>";
            zos.write(documentXml.getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();
        }
        byte[] validDocxBytes = baos.toByteArray(); // resulting byte array (validDocxBytes) now represents a minimal valid DOCX file

        // Mock: Mocks the behavior of mockFile object: getOriginalFilename() and getInputStream()
        when(mockFile.getOriginalFilename()).thenReturn("test.docx");
        when(mockFile.getInputStream()).thenReturn(new ByteArrayInputStream(validDocxBytes));

        // Invoke the controller method that handles the file import.
        String result = documentController.importFile(mockFile, null);

        // Assert: Verify that a non-null result is returned.
        assertNotNull(result);
    }

    @Test
    void testAppendSignature_Success() {
        // Mock: Creates a mock MultipartFile to simulate a file upload
        MultipartFile mockFile = mock(MultipartFile.class);

        // Mock: the behavior of the appendSignature() method in DocumentServiceImpl
        when(documentServiceImpl.appendSignature(mockFile, "Atul")).thenReturn("mocked_sfdt_content");

        // Invoke
        String response = documentController.appendSignature(mockFile, "Atul");

        // Assert: Verify that a non-null response is returned
        assertNotNull(response);
    }

    @Test
    void testSaveFile_Success() throws Exception {
        // Arrange: Create a dummy SaveDto with a file name and content.
        SaveDto saveDto = new SaveDto();

        // In saveDto: Use a temporary file to avoid overwriting real files.
        String tempFileName = "temp_testDocument.docx";
        saveDto.setFileName(tempFileName);
        saveDto.setContent("dummyContent");

        // Use Mockito's static mocking to stub the WordProcessorHelper.save method.
        //Mockito's static mocking (with try-with-resources) for WordProcessorHelper.save(String).
        // When called with "dummyContent", it returns a mocked WordDocument

        try (MockedStatic<WordProcessorHelper> mockedHelper = Mockito.mockStatic(WordProcessorHelper.class)) {
            // Create a mock WordDocument instance.
            WordDocument mockDoc = mock(WordDocument.class);

            // Stub the static method call to return the mock document.
            mockedHelper.when(() -> WordProcessorHelper.save("dummyContent"))
                    .thenReturn(mockDoc);

            // Act: Call the controller's saveFile method.
            documentController.saveFile(saveDto);

            // Assert:
            // Verify that the mock document's save method was called with an OutputStream and a FormatType.
            verify(mockDoc, times(1)).save(any(OutputStream.class), any(FormatType.class));
            // Verify that the document's close method was called.
            verify(mockDoc, times(1)).close();

            // Optionally: Check that the file was actually created.
            File file = new File(tempFileName);
            assertTrue(file.exists(), "File should exist after saving");


            // To see the saved test file comment out the below line
            file.delete(); // Clean up the file after test.
        }
    }

}
