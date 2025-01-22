package org.cas.inlinedocumentmgmtservice;

import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import com.syncfusion.licensing.*;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import java.io.File;
import java.io.InputStream;

@SpringBootApplication
public class InlineDocumentMgmtServiceApplication {

    public static void main(String[] args) {
        SpringApplication.run(InlineDocumentMgmtServiceApplication.class, args);

        // Register Syncfusion license
        SyncfusionLicenseProvider.registerLicense("GTIlMmhha31ifWBgaGBifGJhfGpqampzYWBpZmppZmpoMicmPxMwPzIhOicqICogJzY+IDo9MH0wPD4=");


        try {
            // 1.
            // Get the downloads folder path
            String downloadsFolderPath = System.getProperty("user.home") + File.separator + "Downloads";


            //Creates an instance of WordDocument Instance (Empty Word Document).
                    WordDocument document = new WordDocument();

            //Add a section and paragraph in the empty document.
                    document.ensureMinimal();
            //Append text to the last paragraph of the document.
                    document.getLastParagraph().appendText("Hello World");
            //Save and close the Word document.
                    document.save(downloadsFolderPath + File.separator + "Result3.docx");

            // 2.
            //Appends merge field to the last paragraph.
            document.getLastParagraph().appendField("FullName", FieldType.FieldMergeField);

            //Saves the Word document.
            document.save(downloadsFolderPath + File.separator + "Template.docx", FormatType.Docx);

            document.close();

            DocumentServiceImpl documentService = new DocumentServiceImpl();
            PlantDto plantDto = new PlantDto();
            plantDto.setPlant("Plant 1 in NERC");

            documentService.mailMerge(plantDto);
        }

        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

}
