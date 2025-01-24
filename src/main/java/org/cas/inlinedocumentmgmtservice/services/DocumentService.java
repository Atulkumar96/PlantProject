package org.cas.inlinedocumentmgmtservice.services;

import com.syncfusion.docio.WordDocument;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDetails;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;

import java.util.Map;

public interface DocumentService {
    public WordDocument importDocument();
    public void mailMerge(PlantDto plantDto);
    public void fetchPlantDetails(String plant);
}
