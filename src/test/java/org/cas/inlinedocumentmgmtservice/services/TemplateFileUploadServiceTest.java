package org.cas.inlinedocumentmgmtservice.services;

import org.cas.inlinedocumentmgmtservice.exceptions.DocumentProcessingException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.MockedStatic;
import org.springframework.boot.web.client.RestTemplateBuilder;
import org.springframework.http.ResponseEntity;
import org.springframework.test.util.ReflectionTestUtils;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

public class TemplateFileUploadServiceTest {
    private TemplateFileUploadService templateFileUploadService;
    private RestTemplate restTemplate;
    private RestTemplateBuilder restTemplateBuilder;

    @BeforeEach
    public void setup() {
        // Create a mocked RestTemplate instance
        restTemplate = mock(RestTemplate.class);

        // Create a mocked RestTemplateBuilder
        restTemplateBuilder = mock(RestTemplateBuilder.class);

        // When build() is called, return our mocked RestTemplate
        when(restTemplateBuilder.build()).thenReturn(restTemplate);

        // Instantiate our service with the mocked RestTemplateBuilder
        templateFileUploadService = new TemplateFileUploadService(restTemplateBuilder);

        // Set the FIREBASE_UPLOAD_URL (normally injected via @Value) to a dummy URL for testing
        ReflectionTestUtils.setField(templateFileUploadService, "FIREBASE_UPLOAD_URL", "http://dummy-upload-url.com");
    }

    @Test
    public void testUploadFile_success() throws Exception {
        // Create a mocked MultipartFile to simulate the input file
        MultipartFile mockFile = mock(MultipartFile.class);
        String fileName = "testFile.docx";
        byte[] dummyContent = "dummy content".getBytes();

        // When getBytes() is called on the file, return our dummy content
        when(mockFile.getBytes()).thenReturn(dummyContent);

        // Simulate a successful POST request by stubbing RestTemplate's postForEntity method
        String expectedResponse = "http://uploaded-file-url.com";
        when(restTemplate.postForEntity(anyString(), any(), eq(String.class)))
                .thenReturn(ResponseEntity.ok(expectedResponse));

        // Call the uploadFile method and capture the returned URL
        String actualResponse = templateFileUploadService.uploadFile(mockFile, fileName);

        // Assert that the returned URL matches our expected response
        assertEquals(expectedResponse, actualResponse);
        // Verify that the file's getBytes() method was indeed called
        verify(mockFile).getBytes();
    }

    /**
     * this test confirms that if an IOException occurs while reading the file,
     * then the service correctly wraps and rethrows it as a DocumentProcessingException
     * with an error message that includes the file name
     * @throws Exception
     */

    @Test
    public void testUploadFile_IOException() throws Exception {
        // MultipartFile is mocked to simulate a file input
        MultipartFile mockFile = mock(MultipartFile.class);
        String fileName = "testFile.docx";

        // Stub the getBytes() method to throw an IOException to simulate an error while reading the file
        // Simulates a scenario where reading the file fails, such as due to a corrupted file or disk error
        when(mockFile.getBytes()).thenThrow(new IOException("Simulated IO error"));

        // Expect a DocumentProcessingException when calling uploadFile due to the IO error
        // This is expected behavior because the service method catches the IOException and wraps it in a DocumentProcessingException
        Exception exception = assertThrows(DocumentProcessingException.class, () -> {
            templateFileUploadService.uploadFile(mockFile, fileName);
        });

        // Optionally, verify that the exception message contains the file name
        // This ensures that the error message is informative, indicating which file caused the problem.
        assertTrue(exception.getMessage().contains(fileName));
    }

    /**
     * This test verifies that the uploadFileAfterProtecting method correctly processes the file
     * While the external HTTP call is stubbed to isolate the test
     * we're ensuring that all the internal steps (document protection, request preparation, and response handling)
     * work correctly together to produce the expected outcome
     * This isolation is a core principle of unit testing
     * @throws Exception
     */

    @Test
    public void testUploadFileAfterProtecting_success() throws Exception {
        // A mocked MultipartFile (named mockFile) is created to simulate an input file
        // A test file name (protectedTestFile.docx) is defined
        // A dummy byte array (protectedContent) is prepared to represent the processed (or "protected") content of the file
        MultipartFile mockFile = mock(MultipartFile.class);
        String fileName = "protectedTestFile.docx";
        byte[] protectedContent = "protected content".getBytes();

        // The service method uploadFileAfterProtecting internally calls a static method DocumentServiceImpl.protectDocument(mockFile) to process the file
        // We use Mockito's static mocking capability (using a try-with-resources block) to simulate this behavior
        // Inside this block, we instruct the static method to return our dummy protectedContent when it's called with mockFile
        try (MockedStatic<DocumentServiceImpl> mockedStatic = mockStatic(DocumentServiceImpl.class)) {
            // Stub the static method to return our dummy protected content when called with our mock file
            mockedStatic.when(() -> DocumentServiceImpl.protectDocument(mockFile))
                    .thenReturn(protectedContent);

            // Simulate a successful POST request by stubbing RestTemplate's postForEntity method
            String expectedResponse = "http://uploaded-protected-file-url.com";
            when(restTemplate.postForEntity(anyString(), any(), eq(String.class)))
                    .thenReturn(ResponseEntity.ok(expectedResponse));

            // The method under test, uploadFileAfterProtecting, is called with the mocked file and file name
            // The returned URL is stored in actualResponse
            // Act
            String actualResponse = templateFileUploadService.uploadFileAfterProtecting(mockFile, fileName);

            // Verify that the returned URL matches our expected response as per the stubbed response
            // Assert
            assertEquals(expectedResponse, actualResponse);

            // Verify that the static protectDocument method was called exactly once with our file
            mockedStatic.verify(() -> DocumentServiceImpl.protectDocument(mockFile), times(1));
        }
    }

    @Test
    public void testUploadFileAfterProtecting_throwsException() throws Exception {
        // Create a mocked MultipartFile
        MultipartFile mockFile = mock(MultipartFile.class);
        String fileName = "protectedTestFile.docx";

        // Use static mocking to simulate an exception from the static protectDocument method
        try (MockedStatic<DocumentServiceImpl> mockedStatic = mockStatic(DocumentServiceImpl.class)) {
            // Stub protectDocument to throw a DocumentProcessingException when invoked
            mockedStatic.when(() -> DocumentServiceImpl.protectDocument(mockFile))
                    .thenThrow(new DocumentProcessingException("Error protecting document", new RuntimeException("Underlying error")));

            // Expect a DocumentProcessingException when calling uploadFileAfterProtecting due to the failure in protectDocument
            Exception exception = assertThrows(DocumentProcessingException.class, () -> {
                templateFileUploadService.uploadFileAfterProtecting(mockFile, fileName);
            });

            // Optionally, check that the exception message includes the file name
            assertTrue(exception.getMessage().contains(fileName));

            // Verify that the static protectDocument method was called once
            mockedStatic.verify(() -> DocumentServiceImpl.protectDocument(mockFile), times(1));
        }
    }
}
