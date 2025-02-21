package org.cas.inlinedocumentmgmtservice.services;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import com.syncfusion.docio.*;
import com.syncfusion.ej2.wordprocessor.WordProcessorHelper;
import com.syncfusion.docio.FormFieldType;
//import com.syncfusion.javahelper.drawing.Color;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import com.syncfusion.javahelper.system.io.FileStreamSupport;
import org.cas.inlinedocumentmgmtservice.dtos.CommentDto;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.exceptions.DocumentProcessingException;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

//import com.syncfusion.javahelper.drawing.Color;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.awt.Color;

import java.io.File;
import java.io.InputStream;
import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

@Service
public class DocumentServiceImpl implements DocumentService{
    private static final Logger LOGGER = Logger.getLogger(DocumentServiceImpl.class.getName());

    private String[] mergeFieldNames = null;
    private String[] mergeFieldValues = null;

    // Import the document from the resources folder
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

    /**
     * Performs mail merge operation with the fetched plant details in mergeFieldNames and mergeFieldValues
     * @param plantDto PlantDto object
     */
    @Override
    public void mailMerge(PlantDto plantDto) {

        // Get the download folder path
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

    /**
     * Extracts review comments from the Word document
     * @param documentPath
     * @return JSON formatted String of the extracted comments
     * @throws DocumentProcessingException
     */
    public String extractReviewComments(String documentPath) throws DocumentProcessingException {
        List<CommentDto> commentsList = new ArrayList<>();

        // Ensure document path is valid
        if (documentPath == null || documentPath.isEmpty()) {
            throw new DocumentProcessingException("Document path cannot be null or empty.");
        }

        try (WordDocument document = new WordDocument(documentPath, FormatType.Docx)) {

            // Ensure document contains comments
            if (document.getComments() == null) {
                return "[]"; // Return empty JSON array if no comments exist
            }

            // Iterate through all comments in the document
            for (Object obj : document.getComments()) {
                if (!(obj instanceof WComment)) continue;

                WComment comment = (WComment) obj;
                WCommentFormat format = comment.getFormat();

                // Ensure format is not null
                if (format == null) continue;

                // Extract author, initials, and datetime safely
                String author = Optional.ofNullable(format.getUser()).orElse("Unknown");
                String initials = Optional.ofNullable(format.getUserInitials()).orElse("N/A");
                LocalDateTime commentDateTime = format.getDateTime() != null ? format.getDateTime() : LocalDateTime.MIN;

                // Extract comment text safely
                String commentText = comment.getTextBody() != null ? extractTextFromWTextBody(comment.getTextBody()) : "";

                // Handle reply comments safely
                String parentInfo = "";
                WComment parentComment = (WComment) comment.getAncestor();
                if (parentComment != null && parentComment.getTextBody() != null) {
                    parentInfo = " (Reply to: " + extractTextFromWTextBody(parentComment.getTextBody()) + ")";
                }

                // Add to list
                commentsList.add(new CommentDto(author, initials, commentText, commentDateTime, parentInfo));
            }

        } catch (JsonProcessingException e) {
            throw new DocumentProcessingException("Error processing JSON output", e);
        } catch (IOException e) {
            throw new DocumentProcessingException("Failed to read input document: " + documentPath, e);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Convert list to JSON
        try {
            ObjectMapper objectMapper = new ObjectMapper();
            objectMapper.registerModule(new JavaTimeModule()); // Enable LocalDateTime support
            return objectMapper.writeValueAsString(commentsList);
        } catch (JsonProcessingException e) {
            throw new DocumentProcessingException("Error serializing comments list", e);
        }
    }


    // TODO: Test this overloaded extractReviewComments method, from Frontend Word Editor or any other client
    /**
     * Extracts review comments from the Word document
     * @param file MultipartFile input stream
     * @return JSON formatted String of the extracted comments
     * @throws DocumentProcessingException
     */
    public String extractReviewComments(MultipartFile file) throws DocumentProcessingException {
        List<CommentDto> commentsList = new ArrayList<>();

        // Use try-with-resources to ensure proper resource management
        try (InputStream inputStream = file.getInputStream();
             WordDocument document = new WordDocument(inputStream, FormatType.Docx)) {

            // Check if document has comments
            if (document.getComments() == null) {
                return "[]"; // Return empty JSON array if no comments exist
            }

            // Iterate through all comments in the document
            for (Object obj : document.getComments()) {
                if (!(obj instanceof WComment)) continue;

                WComment comment = (WComment) obj;
                WCommentFormat format = comment.getFormat();

                // Ensure format is not null
                if (format == null) continue;

                // Extract author, initials, and datetime safely
                String author = Optional.ofNullable(format.getUser()).orElse("Unknown");
                String initials = Optional.ofNullable(format.getUserInitials()).orElse("N/A");
                LocalDateTime commentDateTime = format.getDateTime() != null ? format.getDateTime() : LocalDateTime.MIN;

                // Extract comment text safely
                String commentText = comment.getTextBody() != null ? extractTextFromWTextBody(comment.getTextBody()) : "";

                // Handle reply comments safely
                String parentInfo = "";
                WComment parentComment = (WComment) comment.getAncestor();
                if (parentComment != null && parentComment.getTextBody() != null) {
                    parentInfo = " (Reply to: " + extractTextFromWTextBody(parentComment.getTextBody()) + ")";
                }

                // Add to list
                commentsList.add(new CommentDto(author, initials, commentText, commentDateTime, parentInfo));
            }

        } catch (JsonProcessingException e) {
            throw new DocumentProcessingException("Error processing JSON output", e);
        } catch (IOException e) {
            throw new DocumentProcessingException("Failed to read input document", e);
        } catch (Exception e) {
            throw new DocumentProcessingException("Exception:", e);
        }

        // Convert list to JSON
        try {
            ObjectMapper objectMapper = new ObjectMapper();
            objectMapper.registerModule(new JavaTimeModule()); // Enable LocalDateTime support
            return objectMapper.writeValueAsString(commentsList);
        } catch (JsonProcessingException e) {
            throw new DocumentProcessingException("Error serializing comments list", e);
        }
    }

    /**
     * Protects the document and makes the content with only green background editable
     * @param documentPath the path of the input document
     * @param outputPath the path of the output document
     * @return void
     * @throws DocumentProcessingException
     */
    public void protectDocument(String documentPath, String outputPath) throws DocumentProcessingException {
        // Validate input parameters
        if (documentPath == null || documentPath.isEmpty()) {
            throw new DocumentProcessingException("Invalid document path provided.");
        }
        if (outputPath == null || outputPath.isEmpty()) {
            throw new DocumentProcessingException("Invalid output path provided.");
        }

        // Use try-with-resources to ensure the document is closed properly
        try (WordDocument document = new WordDocument(documentPath, FormatType.Docx)) {

            // Iterate through the sections in the document
            if (document.getSections() != null) {
                for (Object sectionObj : document.getSections()) {
                    if (!(sectionObj instanceof WSection)) {
                        continue;
                    }
                    WSection section = (WSection) sectionObj;

                    // Iterate through the child entities in the section
                    if (section.getBody() == null || section.getBody().getChildEntities() == null) {
                        continue;
                    }
                    for (Object entity : section.getBody().getChildEntities()) {
                        // Process tables
                        if (entity instanceof WTable) {
                            WTable table = (WTable) entity;
                            boolean hasGreenRow = false;

                            if (table.getRows() != null) {
                                // Iterate through the rows of the table
                                for (Object rowObj : table.getRows()) {
                                    if (!(rowObj instanceof WTableRow)) {
                                        continue;
                                    }
                                    WTableRow row = (WTableRow) rowObj;

                                    if (row.getCells() != null) {
                                        // Iterate through the cells of the row
                                        for (Object cellObj : row.getCells()) {
                                            if (!(cellObj instanceof WTableCell)) {
                                                continue;
                                            }
                                            WTableCell cell = (WTableCell) cellObj;

                                            // If the cell has a green background then make that cell editable
                                            if (cell.getCellFormat() != null &&
                                                    Objects.equals(cell.getCellFormat().getBackColor(), ColorSupport.fromArgb(-1, -51, -1, -51))) {

                                                hasGreenRow = true; // Mark that this table has editable cells

                                                if (cell.getParagraphs() != null) {
                                                    // Iterate through the paragraphs in the cell
                                                    for (Object paraObj : cell.getParagraphs()) {
                                                        if (!(paraObj instanceof WParagraph)) {
                                                            continue;
                                                        }
                                                        WParagraph paragraphInGreen = (WParagraph) paraObj;

                                                        // Get the text inside the paragraph
                                                        String paragraphText = paragraphInGreen.getText();

                                                        // Clear existing text to replace with an editable content control
                                                        paragraphInGreen.setText("");

                                                        // Create an inline content control (editable field)
                                                        IInlineContentControl contentControl = (InlineContentControl) paragraphInGreen.appendInlineContentControl(ContentControlType.Text);

                                                        // Set the content inside the content control
                                                        WTextRange textRange = new WTextRange(document);
                                                        textRange.setText(paragraphText);
                                                        contentControl.getParagraphItems().add(textRange);

                                                        // Set content control properties
                                                        if (contentControl.getContentControlProperties() != null) {
                                                            contentControl.getContentControlProperties().setLockContentControl(false);
                                                            contentControl.getContentControlProperties().setLockContents(false);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            // Optional: If the table has at least one green-background row,
                            // programmatically add an empty editable row at the end.
                            // if (hasGreenRow) {
                            //     addEditableRowToTable(document, table);
                            // }
                        }
                        // Process paragraphs
                        else if (entity instanceof WParagraph) {
                            WParagraph paragraph = (WParagraph) entity;
                            if (paragraph.getParagraphFormat() != null &&
                                    Objects.equals(paragraph.getParagraphFormat().getBackColor(), ColorSupport.fromArgb(-1, -51, -1, -51))) {

                                // Get the text inside the paragraph
                                String paragraphValue = paragraph.getText();

                                // Clear existing text to replace with an editable content control
                                paragraph.setText("");

                                // Create an inline content control (editable field)
                                IInlineContentControl contentControl = (InlineContentControl) paragraph.appendInlineContentControl(ContentControlType.Text);

                                // Set the content inside the content control
                                WTextRange textRange = new WTextRange(document);
                                textRange.setText(paragraphValue);
                                contentControl.getParagraphItems().add(textRange);

                                // Set content control properties
                                if (contentControl.getContentControlProperties() != null) {
                                    contentControl.getContentControlProperties().setLockContentControl(false);
                                    contentControl.getContentControlProperties().setLockContents(false);
                                }
                            }
                        }
                    }
                }
            }

            // Set document protection to allow only form fields modifications
            document.protect(ProtectionType.AllowOnlyFormFields);

            // Save the protected document to the specified output path
            document.save(outputPath);

        } catch (Exception e) {
            // Wrap any exception in a custom exception for higher-level handling
            throw new DocumentProcessingException("Failed to protect document.", e);
        }
    }

    // TODO: Test this overloaded protectDocument method, from Frontend Word Editor or any other client
    /**
     * Protects the document and makes the content with only green background editable
     * @param file MultipartFile input stream
     * @return void
     * @throws DocumentProcessingException
     */
    public void protectDocument(MultipartFile file) throws DocumentProcessingException {
        // Use try-with-resources to ensure the document is closed properly
        try (InputStream inputStream = file.getInputStream();
             WordDocument document = new WordDocument(inputStream, FormatType.Docx)) {

            // Iterate through the sections in the document
            if (document.getSections() != null) {
                for (Object sectionObj : document.getSections()) {
                    if (!(sectionObj instanceof WSection)) {
                        continue;
                    }
                    WSection section = (WSection) sectionObj;

                    // Iterate through the child entities in the section
                    if (section.getBody() == null || section.getBody().getChildEntities() == null) {
                        continue;
                    }
                    for (Object entity : section.getBody().getChildEntities()) {
                        // Process tables
                        if (entity instanceof WTable) {
                            WTable table = (WTable) entity;
                            boolean hasGreenRow = false;

                            if (table.getRows() != null) {
                                // Iterate through the rows of the table
                                for (Object rowObj : table.getRows()) {
                                    if (!(rowObj instanceof WTableRow)) {
                                        continue;
                                    }
                                    WTableRow row = (WTableRow) rowObj;

                                    if (row.getCells() != null) {
                                        // Iterate through the cells of the row
                                        for (Object cellObj : row.getCells()) {
                                            if (!(cellObj instanceof WTableCell)) {
                                                continue;
                                            }
                                            WTableCell cell = (WTableCell) cellObj;

                                            // If the cell has a green background then make that cell editable
                                            if (cell.getCellFormat() != null &&
                                                    Objects.equals(cell.getCellFormat().getBackColor(), ColorSupport.fromArgb(-1, -51, -1, -51))) {

                                                hasGreenRow = true; // Mark that this table has editable cells

                                                if (cell.getParagraphs() != null) {
                                                    // Iterate through the paragraphs in the cell
                                                    for (Object paraObj : cell.getParagraphs()) {
                                                        if (!(paraObj instanceof WParagraph)) {
                                                            continue;
                                                        }
                                                        WParagraph paragraphInGreen = (WParagraph) paraObj;

                                                        // Get the text inside the paragraph
                                                        String paragraphText = paragraphInGreen.getText();

                                                        // Clear existing text to replace with an editable content control
                                                        paragraphInGreen.setText("");

                                                        // Create an inline content control (editable field)
                                                        IInlineContentControl contentControl = (InlineContentControl) paragraphInGreen.appendInlineContentControl(ContentControlType.Text);

                                                        // Set the content inside the content control
                                                        WTextRange textRange = new WTextRange(document);
                                                        textRange.setText(paragraphText);
                                                        contentControl.getParagraphItems().add(textRange);

                                                        // Set content control properties
                                                        if (contentControl.getContentControlProperties() != null) {
                                                            contentControl.getContentControlProperties().setLockContentControl(false);
                                                            contentControl.getContentControlProperties().setLockContents(false);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            // Optional: If the table has at least one green-background row,
                            // programmatically add an empty editable row at the end.
                            // if (hasGreenRow) {
                            //     addEditableRowToTable(document, table);
                            // }
                        }
                        // Process paragraphs
                        else if (entity instanceof WParagraph) {
                            WParagraph paragraph = (WParagraph) entity;
                            if (paragraph.getParagraphFormat() != null &&
                                    Objects.equals(paragraph.getParagraphFormat().getBackColor(), ColorSupport.fromArgb(-1, -51, -1, -51))) {

                                // Get the text inside the paragraph
                                String paragraphValue = paragraph.getText();

                                // Clear existing text to replace with an editable content control
                                paragraph.setText("");

                                // Create an inline content control (editable field)
                                IInlineContentControl contentControl = (InlineContentControl) paragraph.appendInlineContentControl(ContentControlType.Text);

                                // Set the content inside the content control
                                WTextRange textRange = new WTextRange(document);
                                textRange.setText(paragraphValue);
                                contentControl.getParagraphItems().add(textRange);

                                // Set content control properties
                                if (contentControl.getContentControlProperties() != null) {
                                    contentControl.getContentControlProperties().setLockContentControl(false);
                                    contentControl.getContentControlProperties().setLockContents(false);
                                }
                            }
                        }
                    }
                }
            }

            // Set document protection to allow only form fields modifications
            document.protect(ProtectionType.AllowOnlyFormFields);

            // If you need to persist the changes, call document.save(...) here

        } catch (Exception e) {
            // Wrap any exception in a custom exception for higher-level handling
            throw new DocumentProcessingException("Failed to protect document.", e);
        }
    }

    /**
     * Adds a new editable row to the specified table.
     * It temporarily removes document protection, adds the row, and then re-applies protection.
     *
     * @param document the WordDocument instance
     * @param table the WTable where the new row will be added
     */
    private void addEditableRowToTable(WordDocument document, WTable table) throws Exception {
        // Remove protection temporarily to allow adding a new row
        document.protect(ProtectionType.NoProtection);

        // Create a new row using the document instance
        WTableRow newRow = new WTableRow(document);

        // Use the first row as a template for cell structure (assuming table has at least one row)
        if (table.getRows() != null) {
            WTableRow templateRow = (WTableRow) table.getRows().get(1);

            // Iterate through the cells of the template row
            for (Object cellObj : templateRow.getCells()) {
                WTableCell newCell = new WTableCell(document);
                // Set the cell's background to the green color used for editable cells
                newCell.getCellFormat().setBackColor(ColorSupport.fromArgb(-1, -51, -1, -51));

                // Create a new paragraph in the cell
                WParagraph newParagraph = new WParagraph(document);
                // Append an inline content control to the paragraph (editable field)
                IInlineContentControl newContentControl = (InlineContentControl) newParagraph.appendInlineContentControl(ContentControlType.Text);

                // Create an empty text range for the content control
                WTextRange newTextRange = new WTextRange(document);
                newTextRange.setText("");  // Empty text for user input
                newContentControl.getParagraphItems().add(newTextRange);

                // Configure content control properties: lock control but allow content editing
                newContentControl.getContentControlProperties().setLockContentControl(true);
                newContentControl.getContentControlProperties().setLockContents(false);

                // Add the paragraph with the content control to the new cell
                newCell.getParagraphs().add(newParagraph);
                // Add the new cell to the new row
                newRow.getCells().add(newCell);
            }

            // Add the new row to the table
            table.getRows().add(newRow);
        }

        // Re-apply document protection to allow modifications only in form fields
        document.protect(ProtectionType.AllowOnlyFormFields);
    }

    // TODO: Once the location of image insertion is confirmed, update/refactor the method accordingly
    /**
     * Inserts an image into the document at the location of the specified text.
     * @param inputFilePath the path of the input document
     * @param inputImagePath the path of the image to insert
     * @param outputFilePath the path of the output document
     * @return void
     */
    public void insertImage(String inputFilePath, String inputImagePath, String outputFilePath) {
        // Validate file paths
        if (!isValidFile(inputFilePath) || !isValidFile(inputImagePath)) {
            LOGGER.severe("Invalid input file paths. Ensure the document and image exist.");
            return;
        }

        try (WordDocument document = new WordDocument(inputFilePath);
             FileInputStream imageStream = new FileInputStream(inputImagePath)) {

            boolean imageInserted = false;

            // Iterate through document sections
            for (Object sectionObj : document.getSections()) {
                WSection section = (WSection) sectionObj;

                // Iterate through child entities (paragraphs, tables, etc.)
                for (Object entity : section.getBody().getChildEntities()) {
                    if (entity instanceof WParagraph) {
                        WParagraph paragraph = (WParagraph) entity;

                        // Insert image in the paragraph containing "Subject Matter Experts"
                        if (paragraph.getText().contains("Subject Matter Experts")) {
                            paragraph.appendText("\n"); // New line before image

                            // Insert the image
                            IWPicture picture = paragraph.appendPicture(imageStream);
                            picture.setHeight(150);
                            picture.setWidth(160);

                            imageInserted = true;
                        }
                    }
                }
            }

            if (imageInserted) {
                document.save(outputFilePath);
                LOGGER.info("Image inserted successfully and document saved at: " + outputFilePath);
            } else {
                LOGGER.warning("No matching text found. Image was not inserted.");
            }

        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "Error inserting image into the document", e);
        }
    }

    // Helper method for insertImage() to validate file paths
    private boolean isValidFile(String filePath) {
        return filePath != null && Files.exists(Paths.get(filePath)) && new File(filePath).isFile();
    }

    // TODO: Test this overloaded insertImage method, from Frontend Word Editor or any other client
    /**
     * Inserts an image into a Word document at the end of any paragraph containing
     * the text "Subject Matter Experts" and returns the updated document as an SFDT string.
     *
     * @param file       the original Word document as a MultipartFile
     * @param inputImage the image file as a MultipartFile
     * @return the modified document as an SFDT string, or a JSON error response if processing fails
     */
    public String insertImage(MultipartFile file, MultipartFile inputImage) {
        // Validate input parameters
        if (file == null || file.isEmpty() || inputImage == null || inputImage.isEmpty()) {
            LOGGER.severe("Invalid input files. Ensure the document and image are provided.");
            return buildErrorResponse("Invalid input files. Ensure the document and image are provided.");
        }

        try (InputStream docStream = file.getInputStream();
             WordDocument document = new WordDocument(docStream)) {

            // Read the image bytes once so that we can create a new stream as needed.
            byte[] imageBytes = inputImage.getBytes();
            boolean imageInserted = false;

            // Iterate through the document sections.
            for (Object sectionObj : document.getSections()) {
                WSection section = (WSection) sectionObj;

                // Iterate through child entities (paragraphs, tables, etc.)
                for (Object entity : section.getBody().getChildEntities()) {
                    if (entity instanceof WParagraph) {
                        WParagraph paragraph = (WParagraph) entity;

                        // Check if the paragraph contains "Subject Matter Experts"
                        if (paragraph.getText().contains("Subject Matter Experts")) {
                            paragraph.appendText("\n"); // Append a new line before image insertion

                            // Create a fresh ByteArrayInputStream for the image bytes.
                            try (ByteArrayInputStream imageStream = new ByteArrayInputStream(imageBytes)) {
                                IWPicture picture = paragraph.appendPicture(imageStream);
                                picture.setHeight(150);  // Set image height
                                picture.setWidth(160);   // Set image width
                            }
                            imageInserted = true;
                        }
                    }
                }
            }

            if (!imageInserted) {
                LOGGER.warning("No matching text found. Image was not inserted.");
            } else {
                LOGGER.info("Image inserted successfully into the document.");
            }

            // Convert the modified document to an SFDT string using the candidate method:
            // String load(WordDocument, boolean)
            String sfdtContent = WordProcessorHelper.load(document, true);
            return sfdtContent;
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "Error inserting image into the document", e);
            return buildErrorResponse(e.getMessage());
        }
    }

    // Helper method for insertImage to build a JSON error response
    /**
     * Returns a JSON-formatted error response.
     *
     * @param message the error message
     * @return a JSON string representing the error in SFDT format
     */
    private String buildErrorResponse(String message) {
        // For production use, consider using a JSON library for proper formatting.
        return "{\"sections\":[{\"blocks\":[{\"inlines\":[{\"text\":\"" + message + "\"}]}]}]}";
    }

    public void insertOle() throws Exception {
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        String inputPdfPath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\testPdf.pdf";
        String inputExcelPath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\testExcel.xlsx";
        String inputDocPath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\testDoc.docx";

        WordDocument document = new WordDocument(inputFilePath);

        // Load the file to be embedded (e.g., a PDF file)
        InputStream oleFileStreamPdf = new FileInputStream(inputPdfPath);
        InputStream oleFileStreamExcel = new FileInputStream(inputExcelPath);
        InputStream oleFileStreamDoc = new FileInputStream(inputDocPath);

        // Load the display image for the PDF
        InputStream pdfimageStream = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\pdf-icon.png");
        WPicture pdfPicture = new WPicture(document);
        pdfPicture.loadImage(pdfimageStream);

        // Set icon size (e.g., 50 x 50)
        pdfPicture.setHeight(50);
        pdfPicture.setWidth(50);

        // Load the display image for the Excel
        InputStream excelImageStream = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\excel-icon.png");
        WPicture excelPicture = new WPicture(document);
        excelPicture.loadImage(excelImageStream);

        // Set icon size (e.g., 50 x 50)
        excelPicture.setHeight(50);
        excelPicture.setWidth(50);

        // Load the display image for the Doc
        InputStream docImageStream = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\doc-icon.png");
        WPicture docPicture = new WPicture(document);
        docPicture.loadImage(docImageStream);

        // Set icon size (e.g., 50 x 50)
        docPicture.setHeight(50);
        docPicture.setWidth(50);

        // Iterate through the sections in the document
        for (Object sectionObj : document.getSections()) {
            WSection section = (WSection) sectionObj;

            // Iterate through the child entities in the section, child entities can be tables, paragraphs etc.
            for (Object entity : section.getBody().getChildEntities()) {
                if (entity instanceof WParagraph) {
                    WParagraph paragraph = (WParagraph) entity;

                    // if the paragraph contains the text like regex "Subject Matter Experts" then insert the image at the end of the paragraph
                    if (paragraph.getText().contains("Subject Matter Experts")) {

                        IWTextRange textRange = paragraph.appendText("\n"); // Append a new line

                        // Append the excel file to the paragraph
                        WOleObject oleObjectExcel = paragraph.appendOleObject(oleFileStreamExcel, excelPicture, OleObjectType.ExcelWorksheet);

                        paragraph.appendText("\n"); // Append a new line

                        // Append the Word document to the paragraph
                        WOleObject oleObjectDoc = paragraph.appendOleObject(oleFileStreamDoc, docPicture, OleObjectType.WordDocument);

                        paragraph.appendText("\n"); // Append a new line

                        // Append the PDF file to the paragraph
                        WOleObject oleObjectPdf = paragraph.appendOleObject(oleFileStreamPdf, pdfPicture, OleObjectType.AdobeAcrobatDocument);

                        //oleObject.setProgId("AcroExch.Document");
                        // Set display properties for the OLE object
                        oleObjectExcel.setDisplayAsIcon(true);
                        oleObjectDoc.setDisplayAsIcon(true);
                        oleObjectPdf.setDisplayAsIcon(true);
                    }

                }
            }

        }

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\insertObjectDoc.docx";
        document.save(outputFilePath);

        // Close all streams and the document
        oleFileStreamExcel.close();
        oleFileStreamDoc.close();
        oleFileStreamPdf.close();

        pdfimageStream.close();
        docImageStream.close();
        excelImageStream.close();

        document.close();
    }

    public void insertLink() throws Exception {
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";

        WordDocument document = new WordDocument(inputFilePath);

        // Iterate through the sections in the document
        for (Object sectionObj : document.getSections()) {
            WSection section = (WSection) sectionObj;

            // Iterate through the child entities in the section, child entities can be tables, paragraphs etc.
            for (Object entity : section.getBody().getChildEntities()) {
                if (entity instanceof WParagraph) {
                    WParagraph paragraph = (WParagraph) entity;

                    // if the paragraph contains the text "Subject Matter Experts" then insert the image at the end of the paragraph
                    if (paragraph.getText().contains("Subject Matter Experts")) {

                        IWTextRange textRange = paragraph.appendText("\n");

                        // Append the link to the paragraph
                        paragraph.appendHyperlink("https://www.google.com/", "Google", HyperlinkType.WebLink);
                    }

                }
            }

        }

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\insertLink.docx";
        document.save(outputFilePath);

        document.close();
    }

    /**
     * Appends a signature to the last paragraph of provided Word document.
     *
     * @param inputFilePath  the path of the input Word document
     * @param outputFilePath the path where the updated document will be saved
     * @param approverName   the name of the approver
     */
    public void appendSignature(String inputFilePath, String outputFilePath, String approverName) {
        try {
            // Validate input parameters
            validateInputs(approverName, inputFilePath);

            // Format the current date to MM/dd/yyyy
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
            String formattedDate = LocalDate.now().format(formatter);

            // Open the document in a try-with-resources block
            try (WordDocument document = new WordDocument(inputFilePath)) {

                boolean signatureAdded = false;

                // Iterate through each section in the document
                for (Object sectionObj : document.getSections()) {
                    WSection section = (WSection) sectionObj;
                    IWParagraph lastParagraph = section.getBody().getLastParagraph();

                    if (lastParagraph != null) {
                        String signatureText = "Approved by: " + approverName + "\nSigned Date: " + formattedDate;
                        IWTextRange textRange = lastParagraph.appendText("\n" + signatureText);
                        textRange.getCharacterFormat().setBold(true);
                        textRange.getCharacterFormat().setFontSize(12);
                        signatureAdded = true;
                    }
                }

                if (!signatureAdded) {
                    throw new DocumentProcessingException("No valid paragraph found to append the signature.");
                }

                // Save the modified document
                document.save(outputFilePath);
                LOGGER.info("Signature appended and document saved at: " + outputFilePath);

            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "Error processing document", e);
                throw new DocumentProcessingException("Failed to process the document.", e);
            }
        } catch (DocumentProcessingException e) {
            throw e; // Ensure it's handled by the GlobalExceptionHandler
        } catch (Exception e) {
            throw new DocumentProcessingException("Unexpected error while processing the document.", e);
        }
    }

    /**
     * Appends a signature to the last paragraph of the provided Word document.
     * The signature includes the approver's name and the current date.
     * The modified document is returned as an SFDT string for the frontend Word editor.
     *
     * @param inputFile   the uploaded Word document as a MultipartFile
     * @param approverName the name of the approver
     * @return the modified document as an SFDT string
     * @throws DocumentProcessingException if processing fails
     */
    public String appendSignature(MultipartFile inputFile, String approverName) {
        try {
            // Validate input parameters
            validateInputs(approverName, inputFile);

            // Format the current date to MM/dd/yyyy
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
            String formattedDate = LocalDate.now().format(formatter);

            // Open the document from the MultipartFile InputStream
            try (WordDocument document = new WordDocument(inputFile.getInputStream())) {
                boolean signatureAdded = false;

                // Iterate through each section in the document
                for (Object sectionObj : document.getSections()) {
                    WSection section = (WSection) sectionObj;
                    IWParagraph lastParagraph = section.getBody().getLastParagraph();

                    if (lastParagraph != null) {
                        String signatureText = "Approved by: " + approverName + "\nSigned Date: " + formattedDate;
                        IWTextRange textRange = lastParagraph.appendText("\n" + signatureText);
                        textRange.getCharacterFormat().setBold(true);
                        textRange.getCharacterFormat().setFontSize(12);
                        signatureAdded = true;
                    }
                }

                if (!signatureAdded) {
                    throw new DocumentProcessingException("No valid paragraph found to append the signature.");
                }

                // Convert the modified document to an SFDT string using the candidate method.
                return WordProcessorHelper.load(document, true);
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "Error processing document", e);
                throw new DocumentProcessingException("Failed to process the document.", e);
            }
        } catch (DocumentProcessingException e) {
            throw e; // Handled by the GlobalExceptionHandler
        } catch (Exception e) {
            throw new DocumentProcessingException("Unexpected error while processing the document.", e);
        }
    }

    /**
     * Validates the approver name and input file.
     *
     * @param approverName the name of the approver
     * @param inputFile    the uploaded document as a MultipartFile
     */
    private void validateInputs(String approverName, MultipartFile inputFile) {
        if (approverName == null || approverName.trim().isEmpty()) {
            throw new DocumentProcessingException("Approver name cannot be null or empty.");
        }
        if (inputFile == null || inputFile.isEmpty()) {
            throw new DocumentProcessingException("Input file is null or empty.");
        }
    }

    /**
     * Validates input parameters before processing the document.
     *
     * @param approverName  Name of the approver
     * @param inputFilePath Path of the input file
     */
    // Helper method for appendSignature to validate input parameters
    private void validateInputs(String approverName, String inputFilePath) {
        if (approverName == null || approverName.trim().isEmpty()) {
            throw new DocumentProcessingException("Approver name cannot be null or empty.");
        }

        File inputFile = new File(inputFilePath);
        if (!inputFile.exists() || !inputFile.isFile()) {
            throw new DocumentProcessingException("Input file does not exist: " + inputFilePath);
        }
    }

    /**
     * Append a signature by iterating through each child entity in the doc -> getting the last paragraph
     * @param approverName
     * @throws Exception
     */
    public void appendSignatureAtLast(String approverName) throws Exception{
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        WordDocument document = new WordDocument(inputFilePath);

        // Format the current date to MM/dd/yyyy
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
        String formattedDate = LocalDate.now().format(formatter);

        // Iterate through each section in the document
        for (Object sectionObj : document.getSections()) {
            WSection section = (WSection) sectionObj;

            // Get the child entities of the section
            List<Entity> childEntities = new ArrayList<>();
            for (Object obj : section.getBody().getChildEntities()) {
                childEntities.add((Entity) obj);
            }

            WParagraph lastParagraph = null;

            // Find the last non-footer paragraph
            for (Entity entity : childEntities) {
                if (entity instanceof WParagraph) {
                    lastParagraph = (WParagraph) entity;  // Keep track of the last paragraph
                }
            }

            // Ensure there's a valid paragraph before the footer
            if (lastParagraph != null) {
                // Create signature text
                String signatureText = "Approved by: " + approverName + "\nSigned Date: " +
                        formattedDate;

                // Append the text to the last paragraph
                IWTextRange textRange = lastParagraph.appendText("\n" + signatureText);
                textRange.getCharacterFormat().setBold(true);  // Make it bold
                textRange.getCharacterFormat().setFontSize(12);  // Set font size
            }
        }

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\AppendedSignature.docx";
        document.save(outputFilePath);
        document.close();
    }

    // mailMerge helper method: to fetch plant details based on the plant name
    @Override
    public void fetchPlantDetails(String plant) {
        // Plant details map
        Map<String, String> plantDetails = new HashMap<>();
        plantDetails.put("plant1", plant);

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

    // extractReviewComments helper method: Extracts text from WTextBody
    private static String extractTextFromWTextBody(WTextBody textBody) throws Exception {
        StringBuilder textContent = new StringBuilder();
        for (Object paraObj : textBody.getChildEntities()) {
            if (paraObj instanceof WParagraph) {
                WParagraph paragraph = (WParagraph) paraObj;
                for (Object child : paragraph.getChildEntities()) {
                    if (child instanceof WTextRange) {
                        textContent.append(((WTextRange) child).getText()).append(" ");
                    }
                }
            }
        }
        return textContent.toString().trim();
    }
}
