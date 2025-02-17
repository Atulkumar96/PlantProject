package org.cas.inlinedocumentmgmtservice.services;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import com.syncfusion.docio.*;
import com.syncfusion.docio.FormFieldType;
//import com.syncfusion.javahelper.drawing.Color;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import org.cas.inlinedocumentmgmtservice.dtos.CommentDto;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

//import com.syncfusion.javahelper.drawing.Color;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.*;
import java.awt.Color;

import java.io.File;
import java.io.InputStream;

@Service
public class DocumentServiceImpl implements DocumentService{
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
     * @throws Exception
     */
    public String extractReviewComments(String documentPath) throws Exception {

        // Load the Word document
        WordDocument document = new WordDocument(documentPath, FormatType.Docx);

        // List to store extracted comments
        List<CommentDto> commentsList = new ArrayList<>();

        // Iterate through all comments in the document
        for (Object obj : document.getComments()){
            WComment comment = (WComment) obj;

            // Fetch comment author, initials, and text
            WCommentFormat format = comment.getFormat();
            String author = format.getUser(); // Fetch Author
            String initials = format.getUserInitials(); // Fetch Author initials
            LocalDateTime commentDateTime =  format.getDateTime(); // Fetch comment date and time

            //- Extract comment text
            String commentText = extractTextFromWTextBody(comment.getTextBody());

            //- Handle reply comments by checking their ancestor (parent comment)
            WComment parentComment = (WComment) comment.getAncestor();

            String parentInfo = (parentComment != null)
                    ? " (Reply to: " + extractTextFromWTextBody(parentComment.getTextBody()) + ")"
                    : "";


            // Add extracted comment to the list
            commentsList.add(new CommentDto(author, initials, commentText, commentDateTime, parentInfo));
        }

        // Close the document and return the comments list
        document.close();

        // Return the comments list as a JSON
        ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.registerModule(new JavaTimeModule()); // Enable LocalDateTime support
        return objectMapper.writeValueAsString(commentsList);
        //return commentsList;
    }

    // TODO: Test this overloaded extractReviewComments method, if required from Frontend Word Editor or any other client
    public List<CommentDto> extractReviewComments(MultipartFile file) throws Exception {
        // Load the Word document from the MultipartFile input stream
        InputStream inputStream = file.getInputStream();
        WordDocument document = new WordDocument(inputStream, FormatType.Docx);

        // List to store extracted comments
        List<CommentDto> commentsList = new ArrayList<>();

        // Iterate through all comments in the document
        for (Object obj : document.getComments()) {
            WComment comment = (WComment) obj;

            // Fetch comment author, initials, and text
            WCommentFormat format = comment.getFormat();
            String author = format.getUser(); // Fetch Author
            String initials = format.getUserInitials(); // Fetch Author initials

            // Fetch comment date and time
            LocalDateTime commentDateTime = format.getDateTime();

            // Extract comment text
            String commentText = extractTextFromWTextBody(comment.getTextBody());

            // Handle reply comments by checking their ancestor (parent comment)
            WComment parentComment = (WComment) comment.getAncestor();
            String parentInfo = (parentComment != null)
                    ? " (Reply to: " + extractTextFromWTextBody(parentComment.getTextBody()) + ")"
                    : "";

            // Add extracted comment to the list
            commentsList.add(new CommentDto(author, initials, commentText, commentDateTime, parentInfo));
        }

        // Close the document and return the comments list
        document.close();
        inputStream.close();

        return commentsList;
    }

    /**
    public String protectDocumentEnableSpecificSection(MultipartFile file) throws Exception {
        try (InputStream inputStream = file.getInputStream();
             WordDocument document = new WordDocument(inputStream, FormatType.Docx)) {

            // Apply document-wide read-only protection
            document.protect(ProtectionType.AllowOnlyReading);

            for (Object obj : document.getSections()) {
                if (obj instanceof WSection section) {
                    boolean isEditable = false;

                    for (Object child : section.getBody().getChildEntities()) {
                        if (child instanceof WParagraph paragraph) {
                            ColorSupport backgroundColor = paragraph.getParagraphFormat().getBackColor();
                            String hexColor = (backgroundColor != null) ? backgroundColor.toString() : "";

                            if ("#00FF00".equalsIgnoreCase(hexColor)) {
                                isEditable = true;
                                break; // If any paragraph has a green background, make the section editable
                            }
                        }
                    }

                    if (isEditable) {
                        addEditableContentControl(section);
                    }
                }
            }

            // Convert document to SFDT format
            String sfdtContent;
            try {
                sfdtContent = WordProcessorHelper.load(document);
            } catch (InvocationTargetException e) {
                Throwable cause = e.getCause();
                throw new Exception("Error loading document: " + (cause != null ? cause.getMessage() : e.getMessage()), e);
            }

            return sfdtContent;
        } catch (Exception e) {
            e.printStackTrace();
            return "{\"sections\":[{\"blocks\":[{\"inlines\":[{\"text\":\"" + e.getMessage() + "\"}]}]}]}";
        }
    }

    private void addEditableContentControl(WSection section) throws Exception {
        // Create a block-level content control for the section body
        BlockContentControl contentControl = new BlockContentControl(section.getDocument(), ContentControlType.RichText);

        // Set content control properties
        contentControl.getContentControlProperties().setAppearance(ContentControlAppearance.Tags);
        contentControl.getContentControlProperties().setTag("Editable Section");
        contentControl.getContentControlProperties().setTitle("Editable Content");

        // Unlock the content inside this section
        contentControl.getContentControlProperties().setLockContentControl(true); // Prevent deletion
        contentControl.getContentControlProperties().setLockContents(false); // Allow editing

        // Move all entities from the section body into the content control
        EntityCollection entities = section.getBody().getChildEntities();
        while (entities.getCount() > 0) {
            IEntity entity = entities.get(0);
            entities.remove(entity);
            contentControl.getChildEntities().add(entity); // Move entity to content control
        }

        // Add the editable content control back to the section body
        section.getBody().getChildEntities().add(contentControl);
    }
    */
    /**
    public void protectDocument() throws Exception {
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        WordDocument document = new WordDocument(inputFilePath);

        // Protect the entire document
        document.protect(ProtectionType.AllowOnlyFormFields, "password");

        // Iterate through each section in the document
        for (int i = 0; i < document.getSections().getCount(); i++) {
            WSection section = document.getSections().get(i);

            // Iterate through each paragraph in the section
            for (int j = 0; j < section.getBody().getChildEntities().getCount(); j++) {
                Object entity = section.getBody().getChildEntities().get(j);

                // Check if the entity is a paragraph
                if (entity instanceof WParagraph) {
                    WParagraph paragraph = (WParagraph) entity;

                    // Iterate through each item in the paragraph
                    for (int k = 0; k < paragraph.getChildEntities().getCount(); k++) {
                        Object item = paragraph.getChildEntities().get(k);

                        // Check if the item is a text range
                        if (item instanceof WTextRange) {
                            WTextRange textRange = (WTextRange) item;

                            // Check if the background color is green
                            if (isGreenBackground(textRange.getCharacterFormat().getHighlightColor())) {
                                // Add a bookmark around the green background text
                                String bookmarkName = "EditableText" + i + "_" + j + "_" + k;
                                paragraph.getItems().insert(k, new BookmarkStart(document, bookmarkName));
                                paragraph.getItems().insert(k + 2, new BookmarkEnd(document, bookmarkName));
                            }
                        }
                    }
                }
            }

            // Iterate through tables in the section
            for (int j = 0; j < section.getTables().getCount(); j++) {
                WTable table = section.getTables().get(j);

                for (int k = 0; k < table.getRows().getCount(); k++) {
                    WTableRow row = table.getRows().get(k);

                    for (int l = 0; l < row.getCells().getCount(); l++) {
                        WTableCell cell = row.getCells().get(l);

                        // Iterate through each paragraph in the table cell
                        for (int m = 0; m < cell.getChildEntities().getCount(); m++) {
                            Object entity = cell.getChildEntities().get(m);

                            // Check if the entity is a paragraph
                            if (entity instanceof WParagraph) {
                                WParagraph paragraph = (WParagraph) entity;

                                // Iterate through each item in the paragraph
                                for (int n = 0; n < paragraph.getChildEntities().getCount(); n++) {
                                    Object item = paragraph.getChildEntities().get(n);

                                    // Check if the item is a text range
                                    if (item instanceof WTextRange) {
                                        WTextRange textRange = (WTextRange) item;

                                        // Check if the background color is green
                                        if (isGreenBackground(textRange.getCharacterFormat().getHighlightColor())) {
                                            // Add a bookmark around the green background text
                                            String bookmarkName = "EditableText" + i + "_" + j + "_" + k + "_" + l + "_" + m;
                                            paragraph.getItems().insert(n, new BookmarkStart(document, bookmarkName));
                                            paragraph.getItems().insert(n + 2, new BookmarkEnd(document, bookmarkName));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\ProtectedDocument.docx";
        document.save(outputFilePath);
        document.close();
    }

    private static boolean isGreenBackground(Object highlightColor) {
        // Implement the method to compare the color
        // Assuming that getHighlightColor() returns a ColorSupport object
        ColorSupport green = ColorSupport.getGreen();
        return highlightColor.equals(green);
    }
    */

    public void protectDocument() throws Exception {
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        WordDocument document = new WordDocument(inputFilePath);

        // Iterate through the sections in the document
        for (Object sectionObj : document.getSections()) {
            WSection section = (WSection) sectionObj;

            // Iterate through the Tables in the section
            for (Object tableObj : section.getTables()) {
                WTable table = (WTable) tableObj;

                //Iterate the rows of the table. table -> rows -> cells -> paragraphs
                for (Object row_tempObj : table.getRows()) {
                    WTableRow row = (WTableRow) row_tempObj;

                    //Iterate through the cells of rows.
                    for (Object cell_tempObj : row.getCells()) {
                        WTableCell cell = (WTableCell) cell_tempObj;

                        //If the cell has BackColor as Green then make that cell editable
                        if (Objects.equals(cell.getCellFormat().getBackColor(), ColorSupport.fromArgb(-1,-51,-1,-51))){

                            //Iterate through the paragraphs of the cell
                            for (Object paragraphs : cell.getParagraphs()) {
                                WParagraph paragraphInGreen = (WParagraph) paragraphs;

                                // Get the text inside the paragraphInGreen
                                String paragraphInGreenValue = paragraphInGreen.getText();

                                // Clear existing text to replace with an editable content control
                                paragraphInGreen.setText("");

                                // Create an inline content control (editable field)
                                IInlineContentControl contentControl = (InlineContentControl) paragraphInGreen.appendInlineContentControl(ContentControlType.Text);

                                // Set the content inside the content control
                                WTextRange textRange = new WTextRange(document);
                                textRange.setText(paragraphInGreenValue);
                                contentControl.getParagraphItems().add(textRange);

                                //Enables content control lock. // Users can't remove the content control
                                contentControl.getContentControlProperties().setLockContentControl(true);
                                //Protects the contents of content control. // Users can edit contents
                                contentControl.getContentControlProperties().setLockContents(false);
                            }

                        }
                        // Referring
//                    if (paragraph.getText().contains("Text entry area with Green background:")) {
//                        //cell.getCellFormat().setBackColor(ColorSupport.getGreen());
//                        ColorSupport backColor = cell.getCellFormat().getBackColor();
//                        System.out.println("Color: "+backColor);
//                    }
                        //}
                    }
                }


            }

            // Iterate through the paragraphs of the section
            for (Object paraObj : section.getBody().getChildEntities()) {
                if (paraObj instanceof WParagraph) {
                    WParagraph paragraph = (WParagraph) paraObj;

                    // Check if the paragraph has a green background
                    if (Objects.equals(paragraph.getParagraphFormat().getBackColor(), ColorSupport.fromArgb(-1,-51,-1,-51))) {

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

                        //Enables content control lock. // Users can't remove the content control
                        contentControl.getContentControlProperties().setLockContentControl(true);
                        //Protects the contents of content control. // Users can edit contents
                        contentControl.getContentControlProperties().setLockContents(false);
                    }
                }
            }

        }

//        WSection section = document.getSections().get(0);
//        WTable table = section.getTables().get(2);



        //Set the protection to allow to modify the form fields type
        document.protect(ProtectionType.AllowOnlyFormFields);

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\ProtectedDocument1.docx";
        document.save(outputFilePath);
        document.close();
    }


//    private boolean isGreenBackground(ColorSupport highlightColor) {
//        // Implement the method to compare the color
//        // Assuming that getHighlightColor() returns a ColorSupport object
//        ColorSupport green = ColorSupport.getGreen();
//        return highlightColor.equals(green);
//    }

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
