package org.cas.inlinedocumentmgmtservice.services;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import com.syncfusion.docio.*;
import com.syncfusion.docio.FormFieldType;
//import com.syncfusion.javahelper.drawing.Color;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import com.syncfusion.javahelper.system.io.FileStreamSupport;
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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
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
     * Protects the document and makes the content with only green background editable
     * @param
     * @return JSON formatted String of the extracted comments
     * @throws Exception
     */
    public void protectDocument() throws Exception {
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        WordDocument document = new WordDocument(inputFilePath);

        // Iterate through the sections in the document
        for (Object sectionObj : document.getSections()) {
            WSection section = (WSection) sectionObj;

            // Iterate through the child entities in the section, child entities can be tables, paragraphs etc.
            for (Object entity : section.getBody().getChildEntities()) {
                if (entity instanceof WTable){
                    WTable table = (WTable) entity;

                    //Iterate the rows of the table. table -> rows -> cells -> paragraphs
                    for (Object row_tempObj : table.getRows()) {
                        WTableRow row = (WTableRow) row_tempObj;

                        //Iterate through the cells of rows.
                        for (Object cell_tempObj : row.getCells()) {
                            WTableCell cell = (WTableCell) cell_tempObj;

                            //If the cell has BackColor as Green then make that cell editable
                            if (Objects.equals(cell.getCellFormat().getBackColor(), ColorSupport.fromArgb(-1,-51,-1,-51))){

                                //System.out.println("Cell Background Color: " + cell.getCellFormat().getBackColor());
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
                else if (entity instanceof WParagraph) {
                    WParagraph paragraph = (WParagraph) entity;

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

        //Set the protection to allow to modify the form fields type
        document.protect(ProtectionType.AllowOnlyFormFields);

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\ProtectedDocument2.docx";
        document.save(outputFilePath);
        document.close();
    }

    public void insertImage() throws Exception {
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        String inputImagePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\istockphoto.jpg";

        WordDocument document = new WordDocument(inputFilePath);

        // Load the image from the specified path
        FileInputStream imageStream = new FileInputStream(inputImagePath);

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

                        // Append the image to the paragraph
                        IWPicture picture = paragraph.appendPicture(imageStream);
                        picture.setHeight(150);  // Set the height of the image
                        picture.setWidth(160);  // Set the width of the image
                    }

                }
            }

        }

        // Save the document to the specified path
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\insertImageDoc.docx";
        document.save(outputFilePath);

        // Close all streams and the document
        imageStream.close();
        document.close();
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
     * Appends a signature to the last paragraph of the document
     * @param approverName
     * @throws Exception
     */
    public void appendSignature(String approverName) throws Exception{
        // Load the document from the specified path
        String inputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\RSAW BAL-001-2_2016_v1.docx";
        WordDocument document = new WordDocument(inputFilePath);

        // Format the current date to MM/dd/yyyy
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
        String formattedDate = LocalDate.now().format(formatter);

        // Iterate through each section in the document
        for (Object sectionObj : document.getSections()) {
            WSection section = (WSection) sectionObj;

            // Get the last paragraph in the section
            IWParagraph lastParagraph = section.getBody().getLastParagraph();

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
        String outputFilePath = "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\AppendedSignatureAtLastPara.docx";
        document.save(outputFilePath);
        document.close();
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
