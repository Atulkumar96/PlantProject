package org.cas.inlinedocumentmgmtservice.dtos;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

@NoArgsConstructor
@AllArgsConstructor
public class SaveDto {
    private String _content;
    private String _fileName;
    private String clientId;

    public String getContent() {
        return _content;
    }

    public String setContent(String value) {
        _content = value;
        return value;
    }

    public String getFileName() {
        return _fileName;
    }

    public String setFileName(String value) {
        _fileName = value;
        return value;
    }

    public String getClientId() {
        return clientId;
    }

    public String setClientId(String value) {
        clientId = value;
        return value;
    }
}
