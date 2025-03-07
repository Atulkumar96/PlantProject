package org.cas.inlinedocumentmgmtservice.services;

import org.cas.inlinedocumentmgmtservice.exceptions.DocumentProcessingException;
import org.cas.inlinedocumentmgmtservice.utils.CustomByteArrayResource;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.web.client.RestTemplateBuilder;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

import static org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl.protectDocument;

@Service
public class TemplateFileUploadService {

    @Value("${firebase.upload.url}")
    private String FIREBASE_UPLOAD_URL;
    //private static final String FIREBASE_UPLOAD_URL = "https://sp3test.claritysystemsinc.com/api/v3/upload";
    private final RestTemplate restTemplate;

    public TemplateFileUploadService(RestTemplateBuilder builder) {
        this.restTemplate = builder.build();
    }

    /**
     * Uploads a file to Firebase and returns the URL of the uploaded file.
     * @param file The file to upload
     * @param fileName The name of the file
     * @return The URL of the uploaded file
     */

    public String uploadFile(MultipartFile file, String fileName){
        try{
            // Convert the MultipartFile into a byte array and wrap it in a custom resource
            // that includes the "required filename". This resource will be used in the multipart request.
            CustomByteArrayResource resource = new CustomByteArrayResource(file.getBytes(), fileName);

            // Prepare a multi-part form-data body and add the file resource with key "file"
            // which is expected by the external upload endpoint.
            MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();

            // "file" is the key expected by the upload endpoint
            body.add("file", resource);

            // Set HTTP headers indicating that the content type is multipart/form-data.
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.MULTIPART_FORM_DATA);

            // Create an HttpEntity object that combines the headers and the body for the request.
            HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

            // Send the POST request to the firebase upload URL and capture the response.
            ResponseEntity<String> response = restTemplate.postForEntity(FIREBASE_UPLOAD_URL, requestEntity, String.class);

            // Return the response body which should contain the URL of the uploaded file.
            return response.getBody();

        }
        catch(IOException e){
            // If any I/O error occurs during file processing, wrap it in a DocumentProcessingException
            // with a descriptive message and propagate the exception.
            throw new DocumentProcessingException("Error processing file upload for file: " + fileName, e);
        }
    }

    /**
     * Protects the Document and ploads a file to Firebase and returns the URL of the uploaded file.
     * @param file The file to upload
     * @param fileName The name of the file
     * @return The URL of the uploaded file
     */
    public String uploadFileAfterProtecting(MultipartFile file, String fileName) {
        try {
            // First, process to protect the document with green backcolor editable regions
            byte[] protectedBytes = protectDocument(file);

            // Wrap the processed bytes in the custom resource with the provided filename
            CustomByteArrayResource resource = new CustomByteArrayResource(protectedBytes, fileName);

            // Prepare a multi-part form-data body and add the file resource with key "file"
            MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
            body.add("file", resource);

            // Set HTTP headers indicating that the content type is multipart/form-data
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.MULTIPART_FORM_DATA);

            // Create an HttpEntity object that combines the headers and the body for the request.
            HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

            // Send the POST request to the firebase upload URL and capture the response.
            ResponseEntity<String> response = restTemplate.postForEntity(FIREBASE_UPLOAD_URL, requestEntity, String.class);

            // Return the response body which should contain the URL of the uploaded file.
            return response.getBody();
        } catch (DocumentProcessingException e) {
            throw new DocumentProcessingException("Error processing file upload for file: " + fileName, e);
        }
    }



}
