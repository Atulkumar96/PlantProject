package org.cas.inlinedocumentmgmtservice;

import org.cas.inlinedocumentmgmtservice.dtos.CommentDto;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.services.DocumentService;
import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
// Refer the licensing package
import com.syncfusion.licensing.SyncfusionLicenseProvider;


import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import org.springframework.core.io.DefaultResourceLoader;
import org.springframework.core.io.ResourceLoader;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class InlineDocumentMgmtServiceApplication {

    public static void main(String[] args) throws Exception {
        SpringApplication.run(InlineDocumentMgmtServiceApplication.class, args);

        // Registered Syncfusion license
        SyncfusionLicenseProvider.registerLicense("GTIlMmhha31ifWBgaGBifGJhfGpqampzYWBpZmppZmpoMicmPxMwPzIhOicqICogJzY+IDo9MH0wPD4=");

        // Test the methods in DocumentServiceImpl
        ResourceLoader resourceLoader = new DefaultResourceLoader();

        //1. Test insertLink() method
        try {

            DocumentServiceImpl documentService = new DocumentServiceImpl(resourceLoader);
            //documentService.insertLink();
        } catch (Exception e) {
            e.printStackTrace();
        }

        //1. Test insertImage() method
        try {
            DocumentServiceImpl documentService = new DocumentServiceImpl(resourceLoader);
            documentService.insertImage("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx",
                    "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\istockphoto.jpg",
                    "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\insertImageDoc.docx");
        } catch (Exception e) {
            e.printStackTrace();
        }

        //2. Test insertOle() method
        try {
            DocumentServiceImpl documentService = new DocumentServiceImpl(resourceLoader);

            String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
            String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\insertObjectDoc.docx";

            String inputEmbeddingPdfPath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\testPdf.pdf";
            String inputEmbeddingExcelPath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\testExcel.xlsx";
            String inputEmbeddingDocPath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\testDoc.docx";

            documentService.insertOle(inputFilePath, outputFilePath, inputEmbeddingPdfPath, inputEmbeddingExcelPath, inputEmbeddingDocPath);
        } catch (Exception e) {
            e.printStackTrace();
        }


        //3. Test extractReviewComments method
        try {
            DocumentServiceImpl documentService = new DocumentServiceImpl(resourceLoader);
            //C:\Users\Lenovo\Desktop\Inline Document Service\Test
            String commentsList = documentService.extractReviewComments("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\Plant-NERC-RCP-CIP-002.docx");
            //for (CommentDto comment : commentsList) {
            System.out.println(commentsList);
            //}
        } catch (Exception e) {
            e.printStackTrace();
        }

        //4. Test protectDocument method
        try {
            DocumentServiceImpl documentService = new DocumentServiceImpl(resourceLoader);
            documentService.protectDocument("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\RSAW BAL-001-2_2016_v1.docx",
                    "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\ProtectedDocumentWithGreenEditable.docx");
        } catch (Exception e) {
            e.printStackTrace();
        }

        //5. Test the appendSignature method
        try {
            DocumentServiceImpl documentService = new DocumentServiceImpl(resourceLoader);
            documentService.appendSignature("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\RSAW BAL-001-2_2016_v1.docx",
            //documentService.appendSignature("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\Plant-NERC-RCP-CIP-002.docx",
                     "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\SignedDocumentTest.docx",
                    "Atul");

            //documentService.appendSignatureAtLast("Atul");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
