package org.cas.inlinedocumentmgmtservice.dtos;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.springframework.http.HttpStatus;

@NoArgsConstructor
@AllArgsConstructor
public class ExceptionDto {
    private String message;
    private String details;

    // Getters and Setters
    public String getDetails() {
        return details;
    }

    public void setDetails(String details) {
        this.details = details;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public String toString() {
        return "ExceptionDto(details=" + this.getDetails() + ", message=" + this.getMessage() + ")";
    }
}
