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
import java.util.HashMap;
import java.util.Map;

@Service
public class DocumentServiceImpl implements DocumentService{
    private String[] mergeFieldNames = null;
    private String[] mergeFieldValues = null;

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

        //1. Import the document
        WordDocument document = importDocument();

        //2. Fetch the plant details based on the plant name
        fetchPlantDetails(plantDto.getPlant());

        //3. Perform mail merge operation with the fetched plant details in mergeFieldNames and mergeFieldValues
        try {
            //Executes the Mail merge operation that replaces the matching field names with field values respectively
            document.getMailMerge().execute(mergeFieldNames, mergeFieldValues);

            //3. Save and close the WordDocument instance
            document.save(downloadsFolderPath + File.separator + "MailMergeWord2.docx");
            document.close();
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    @Override
    public void fetchPlantDetails(String plant) {
        // Plant details map
        Map<String, String> plantDetails = new HashMap<>();
        plantDetails.put("plant", plant);

        // TODO Elasticsearch query to fetch plant details based on plant name

        // Initializing the merge field names and values arrays based on the plant details map size
        mergeFieldNames = new String[plantDetails.size()];
        mergeFieldValues = new String[plantDetails.size()];

        // Populate the merge field names and values from the plant details map
        int i = 0;
        for (Map.Entry<String, String> entry : plantDetails.entrySet()) {
            mergeFieldNames[i] = entry.getKey();
            mergeFieldValues[i] = entry.getValue();
            i++;
        }
    }
}
