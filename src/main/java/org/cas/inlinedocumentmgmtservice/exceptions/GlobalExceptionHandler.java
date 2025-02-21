package org.cas.inlinedocumentmgmtservice.exceptions;

import org.cas.inlinedocumentmgmtservice.dtos.ExceptionDto;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

@RestControllerAdvice
public class GlobalExceptionHandler {

    @ExceptionHandler(DocumentProcessingException.class)
    public ResponseEntity<ExceptionDto> handleDocumentProcessingException(DocumentProcessingException ex) {
        return ResponseEntity.status(HttpStatus.BAD_REQUEST) // 400 Bad Request
                .body(new ExceptionDto(ex.getMessage(), "Check logs for more details"));
    }
}
