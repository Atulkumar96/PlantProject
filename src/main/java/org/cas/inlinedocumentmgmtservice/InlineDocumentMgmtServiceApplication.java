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
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class InlineDocumentMgmtServiceApplication {

    public static void main(String[] args) {
        SpringApplication.run(InlineDocumentMgmtServiceApplication.class, args);

        // Registered Syncfusion license
        SyncfusionLicenseProvider.registerLicense("GTIlMmhha31ifWBgaGBifGJhfGpqampzYWBpZmppZmpoMicmPxMwPzIhOicqICogJzY+IDo9MH0wPD4=");

        //1. Test extractReviewComments method
        try {
            DocumentServiceImpl documentService = new DocumentServiceImpl();
            //C:\Users\Lenovo\Desktop\Inline Document Service\Test
            String commentsList = documentService.extractReviewComments("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\Plant-NERC-RCP-CIP-002.docx");
            //for (CommentDto comment : commentsList) {
            System.out.println(commentsList);
            //}
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
