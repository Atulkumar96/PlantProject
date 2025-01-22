package org.cas.inlinedocumentmgmtservice.services;

import com.syncfusion.docio.WordDocument;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDetails;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;

public interface DocumentService {
    public WordDocument importDocument();
    public void mailMerge(PlantDto plantDto);
    public PlantDetails fetchPlantDetails(PlantDto plantDto);
}
