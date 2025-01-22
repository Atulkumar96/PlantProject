package org.cas.inlinedocumentmgmtservice.services;

import com.syncfusion.docio.MailMergeDataTable;
import com.syncfusion.docio.WordDocument;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDetails;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

@Service
public class DocumentServiceImpl implements DocumentService{

    @Override
    public WordDocument importDocument() {

        try {
            // Read the document from the resources folder
            ClassPathResource classPathResource = new ClassPathResource("Plant-NERC-RCP-CIP-002.docx");

            // Get the input stream from the class path resource
            InputStream inputStream = classPathResource.getInputStream();

            // Return the WordDocument instance from the input stream
            return new WordDocument(inputStream);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void mailMerge(PlantDto plantDto) {

        //.. Get the downloads folder path
        String downloadsFolderPath = System.getProperty("user.home") + File.separator + "Downloads";

        // Import the document
        WordDocument document = importDocument();

        // Perform mail merge

        //Initialize the string array with merge field names
        //String[] mergeFieldNames = new String[]{"plant", "plantManagerName", "cipSeniorManagerName", "plantAcronym", "plantCapacity", "city", "state", "reName", "reAcronym"};
        String[] mergeFieldNames = new String[]{"plant"};

        //Initialize the string array with field values
        String[] mergeFieldValues = new String[]{plantDto.getPlant()};

        try {
            //Executes the Mail merge operation that replaces the matching field names with field values respectively
            document.getMailMerge().execute(mergeFieldNames, mergeFieldValues);

            //Save and close the WordDocument instance
            document.save(downloadsFolderPath + File.separator + "MailMergeWord.docx");
            document.close();
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    @Override
    public PlantDetails fetchPlantDetails(PlantDto plantDto) {
        // TODO Elasticsearch query to fetch plant details based on plantDto
        return null;
    }
}
