package org.cas.inlinedocumentmgmtservice.dtos;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Setter
@Getter
@NoArgsConstructor
@AllArgsConstructor
public class PlantDto {
    private int id;
    public static String plant;

    // Setter for plant
    public void setPlant(String plant) {
        this.plant = plant;
    }

    // Getter for plant
    public String getPlant() {
        return plant;
    }

}
