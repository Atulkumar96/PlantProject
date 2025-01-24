package org.cas.inlinedocumentmgmtservice.controllers;

import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.dtos.ResponseDto;
import org.cas.inlinedocumentmgmtservice.services.DocumentService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/documents")
public class DocumentController {
    private final DocumentService documentService;

    public DocumentController(DocumentService documentService) {
        this.documentService = documentService;
    }

    /**
     * This endpoint is used to generate a document by mail merge operation.
     * @param plantDto
     * @return
     * Todo: On the basis of standard- import the document
     * standard: NUMBER;
     * client id: plantNameId is the plant name according to which plant details will be fetched for mail merge
     */

    @PostMapping("/generate")
    public ResponseEntity<ResponseDto> mailMerge(@RequestBody PlantDto plantDto) {
        documentService.mailMerge(plantDto);
        return ResponseEntity.ok(new ResponseDto(HttpStatus.OK, "Mail merge successful"));
    }
}
