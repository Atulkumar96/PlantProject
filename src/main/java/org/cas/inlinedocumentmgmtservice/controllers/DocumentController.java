package org.cas.inlinedocumentmgmtservice.controllers;

import com.syncfusion.docio.WordDocument;
import com.syncfusion.javahelper.system.io.StreamSupport;
import com.syncfusion.javahelper.system.reflection.AssemblySupport;
import com.twelvemonkeys.imageio.plugins.tiff.TIFFImageReaderSpi;
import org.apache.commons.io.IOUtils;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.dtos.ResponseDto;
import org.cas.inlinedocumentmgmtservice.dtos.SaveDto;
import org.cas.inlinedocumentmgmtservice.exceptions.DocumentProcessingException;
import org.cas.inlinedocumentmgmtservice.services.DocumentService;
import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;
import org.cas.inlinedocumentmgmtservice.services.TemplateFileUploadService;
import org.cas.inlinedocumentmgmtservice.utils.CustomByteArrayResource;
import org.springframework.http.*;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.ImageWriter;
import javax.imageio.spi.IIORegistry;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import com.syncfusion.ej2.wordprocessor.WordProcessorHelper;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.ej2.wordprocessor.FormatType;

import com.syncfusion.ej2.wordprocessor.MetafileImageParsedEventArgs;
import com.syncfusion.ej2.wordprocessor.MetafileImageParsedEventHandler;

@CrossOrigin(origins = "*", allowedHeaders = "*")
@RestController
@RequestMapping("/api/inlineDocumentService")
public class DocumentController {
    private final DocumentService documentService;
    private final DocumentServiceImpl documentServiceImpl;
    private final TemplateFileUploadService templateFileUploadService;

    public DocumentController(DocumentService documentService, DocumentServiceImpl documentServiceImpl, TemplateFileUploadService templateFileUploadService) {
        this.documentService = documentService;
        this.documentServiceImpl = documentServiceImpl;
        this.templateFileUploadService = templateFileUploadService;
    }

    /**
     * This endpoint is used to generate a document by mail merge operation.
     * @return
     * Todo: On the basis of standard- import the document
     * Todo: On the basis of plantName- fetch the plant details
     * Todo: When import with no file then fetch the doc. from Firebase and do the generation/mailmerge operation
     * standard: NUMBER;
     * client id: plantNameId is the plant name according to which plant details will be fetched for mail merge
     */

    @PostMapping("/generate")
    public ResponseEntity<ResponseDto> mailMerge(@RequestParam String plantName) {
        documentService.mailMerge(plantName);
        return ResponseEntity.ok(new ResponseDto(HttpStatus.OK, "Mail merge successful"));
    }

    @PostMapping("/policyANDprocedure")
    public ResponseEntity<String> uploadPolicyAndProcedure(
            @RequestParam("clientId") String clientId,
            @RequestParam("standard") String standard,
            @RequestParam("file") MultipartFile file) {
        try {
            String fileName;
            if (clientId == null || clientId.trim().isEmpty()) {
                fileName = standard + "_policyANDprocedure.docx";
            } else {
                fileName = clientId + "_" + standard + "_policyANDprocedure.docx";
            }

            String uploadedUrl = templateFileUploadService.uploadFile(file, fileName);
            return ResponseEntity.ok(uploadedUrl);
        }
        catch (DocumentProcessingException e){
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("File upload failed: " + e.getMessage());
        }
    }


    @PostMapping("/RSAW")
    public ResponseEntity<String> uploadRSAW(
            @RequestParam("clientId") String clientId,
            @RequestParam("standard") String standard,
            @RequestParam("file") MultipartFile file) {
        try {
            String fileName;
            if (clientId == null || clientId.trim().isEmpty()) {
                fileName = standard + "_RSAW.docx";
            } else {
                fileName = clientId + "_" + standard + "_RSAW.docx";
            }

            String uploadedUrl = templateFileUploadService.uploadFileAfterProtecting(file, fileName);
            return ResponseEntity.ok(uploadedUrl);
        }
        catch (DocumentProcessingException e){
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("File upload failed: " + e.getMessage());
        }
    }


    @PostMapping("/appendSignature")
    public String appendSignature(@RequestParam("files") MultipartFile file,
                                  @RequestParam("approverName") String approverName) {
        return documentServiceImpl.appendSignature(file, approverName);
    }

    /**
     * This endpoint is used to import a file and return the content in Sfdt format
     * If no file is provided, the plantName is used to AutoGenerate the document by performing mail merge operation
     * @param file
     * @param plantName
     * @return
     * @throws Exception
     */

    //@CrossOrigin(origins = "*", allowedHeaders = "*")
    @PostMapping("Import")
        public String importFile(@RequestPart(value = "files", required = false) MultipartFile file,
                                 @RequestParam(value = "plantName", required = false) String plantName) throws Exception {

        // If no file is provided, use the PlantDto to perform mail merge
        if (file == null || file.isEmpty()) {
            if (plantName == null) {
                throw new Exception("Either import a file or provide plant details to AutoGenerate the document");
            }

            // Perform mail merge operation to generate the document
            documentService.mailMerge(plantName);
            return "Mail merge successful";
        }

        try {
            String name = file.getOriginalFilename();
            if (name == null || name.isEmpty()) {
                name = "NewDocument1.docx";
            }
            String format = retrieveFileType(name);

            MetafileImageParsedEventHandler metafileImageParsedEvent = new MetafileImageParsedEventHandler() {

                ListSupport<MetafileImageParsedEventHandler> delegateList = new ListSupport<MetafileImageParsedEventHandler>(
                        MetafileImageParsedEventHandler.class);

                // Represents event handling for MetafileImageParsedEventHandlerCollection.
                public void invoke(Object sender, MetafileImageParsedEventArgs args) throws Exception {
                    onMetafileImageParsed(sender, args);
                }

                // Represents the method that handles MetafileImageParsed event.
                public void dynamicInvoke(Object... args) throws Exception {
                    onMetafileImageParsed((Object) args[0], (MetafileImageParsedEventArgs) args[1]);
                }

                // Represents the method that handles MetafileImageParsed event to add collection item.
                public void add(MetafileImageParsedEventHandler delegate) throws Exception {
                    if (delegate != null)
                        delegateList.add(delegate);
                }

                // Represents the method that handles MetafileImageParsed event to remove collection
                // item.
                public void remove(MetafileImageParsedEventHandler delegate) throws Exception {
                    if (delegate != null)
                        delegateList.remove(delegate);
                }
            };
            // Hooks MetafileImageParsed event.
            WordProcessorHelper.MetafileImageParsed.add("OnMetafileImageParsed", metafileImageParsedEvent);
            // Converts DocIO DOM to SFDT DOM.
            //String sfdtContent = WordProcessorHelper.load(file.getInputStream(), getFormatType(format));

            String sfdtContent;

            InputStream inputStream;
            if (".docx".equalsIgnoreCase(format) || ".docm".equalsIgnoreCase(format) ||
                    ".dotx".equalsIgnoreCase(format) || ".dotm".equalsIgnoreCase(format)) {
                inputStream = sanitizeDocxInputStream(file.getInputStream());
            } else {
                inputStream = file.getInputStream();
            }

            try {
                sfdtContent = WordProcessorHelper.load(inputStream, getFormatType(format));
                //sfdtContent = WordProcessorHelper.load(file.getInputStream(), getFormatType(format));
            } catch (InvocationTargetException e) {
                Throwable cause = e.getCause();
                throw new Exception("Error loading document: " + (cause != null ? cause.getMessage() : e.getMessage()), e);
            }

            // Unhooks MetafileImageParsed event.
            WordProcessorHelper.MetafileImageParsed.remove("OnMetafileImageParsed", metafileImageParsedEvent);
            return sfdtContent;
        } catch (Exception e) {
            e.printStackTrace();
            return "{\"sections\":[{\"blocks\":[{\"inlines\":[{\"text\":" + e.getMessage() + "}]}]}]}";
        }
    }

    /**
     * This endpoint is used to save the document to firebase as clientId_Filename.docx and return the link
     * Todo: Save the document to firebase as clientId_Filename.docx and return the link
     * @param data
     * @throws Exception
     */

    @PostMapping("Save")
    public ResponseEntity<String> saveFile(@RequestBody SaveDto data) throws Exception {
        try {
            // Validate that clientId is provided (ensure SaveDto has a clientId field)
            if (data.getClientId() == null || data.getClientId().trim().isEmpty()) {
                throw new Exception("Client ID is required to save the document");
            }

            // Use provided file name or default if not available
            String originalFileName = data.getFileName();
            if (originalFileName == null || originalFileName.isEmpty()) {
                originalFileName = "Document1.docx";
            }

            // Prefix the file name with clientId
            String newFileName = data.getClientId() + "_" + originalFileName;

            // 1. Create the WordDocument in memory using Syncfusion helper from the provided content
            WordDocument document = WordProcessorHelper.save(data.getContent());

            // 2. Write the document to a ByteArrayOutputStream instead of a temporary file
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            String format = retrieveFileType(newFileName); // e.g., ".docx"
            document.save(baos, getWFormatType(format));
            document.close();

            // Convert the output stream to a byte array and wrap it in the CustomByteArrayResource
            CustomByteArrayResource resource = new CustomByteArrayResource(baos.toByteArray(), newFileName);

            // Prepare a multi-part form-data request body using the CustomByteArrayResource
            MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
            body.add("file", resource);

            // Set up HTTP headers for multipart/form-data
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.MULTIPART_FORM_DATA);

            HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

            return ResponseEntity.ok(templateFileUploadService.uploadFileToFirebaseForSave(requestEntity));
        }
        catch (Exception ex) {
            throw new Exception("Error saving file: " + ex.getMessage());
        }
    }

    // The below saveFile method saves the document to the local file system.
    //@CrossOrigin(origins = "*", allowedHeaders = "*")
    @PostMapping("SaveToLocal")
    public void saveFileToLocal(@RequestBody SaveDto data) throws Exception {
        try {
            String name = data.getFileName();
            String format = retrieveFileType(name);
            if (name == null || name.isEmpty()) {
                name = "Document1.docx";
            }
            WordDocument document = WordProcessorHelper.save(data.getContent());
            FileOutputStream fileStream = new FileOutputStream(name);
            document.save(fileStream, getWFormatType(format));
            fileStream.close();
            document.close();
        } catch (Exception ex) {
            throw new Exception(ex.getMessage());
        }
    }


    /**
    @PostMapping(value = "/protectDocument", consumes = MediaType.MULTIPART_FORM_DATA_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<String> protectDocumentSections(@RequestParam("file") MultipartFile file) {
        try {
            String sfdtContent = documentServiceImpl.protectDocumentEnableSpecificSection(file);

            convertSfdtToDocx(sfdtContent, "C:\\Users\\Lenovo\\Desktop\\Inline Document Service\\Test\\ProtectedDocument.docx");

            return ResponseEntity.ok(sfdtContent);
        } catch (Exception e) {
            return ResponseEntity.internalServerError().body("{\"error\":\"" + e.getMessage() + "\"}");
        }
    }
    */

    // importFile helper method
    // Helper method for appendSignature to sanitize the DOCX InputStream
    private InputStream sanitizeDocxInputStream(InputStream docxInputStream) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        // Use try-with-resources for both Zip streams.
        try (ZipInputStream zis = new ZipInputStream(docxInputStream);
             ZipOutputStream zos = new ZipOutputStream(baos)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                String entryName = entry.getName();
                // Create a new entry in the output zip with the same name
                zos.putNextEntry(new ZipEntry(entryName));
                if (entryName.endsWith(".xml")) {
                    // Read the XML content as a String (assume UTF-8 encoding)
                    String xmlContent = IOUtils.toString(zis, StandardCharsets.UTF_8);
                    // Sanitize the XML content by removing invalid XML characters (e.g. U+0002)
                    String sanitizedXml = sanitizeXmlContent(xmlContent);
                    zos.write(sanitizedXml.getBytes(StandardCharsets.UTF_8));
                } else {
                    // For non-XML entries, simply copy the raw bytes.
                    IOUtils.copy(zis, zos);
                }
                zos.closeEntry();
            }
        }
        return new ByteArrayInputStream(baos.toByteArray());
    }

    /**
     * Removes invalid XML characters from a String.
     * The regex below removes control characters except for allowed whitespace (tab, newline, carriage return).
     *
     * @param xml The XML content to sanitize.
     * @return The sanitized XML content.
     */
    // Helper method for importFile to remove invalid XML characters
    private String sanitizeXmlContent(String xml) {
        // Remove control characters in the range 0x00-0x08, 0x0B-0x0C, and 0x0E-0x1F.
        return xml.replaceAll("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]", "");
    }

    // importFile helper method
    private String retrieveFileType(String name) {
        int index = name.lastIndexOf('.');
        String format = index > -1 && index < name.length() - 1 ? name.substring(index) : ".docx";
        return format;
    }

    // importFile helper method
    // Converts Metafile to raster image.
    private static void onMetafileImageParsed(Object sender, MetafileImageParsedEventArgs args) throws Exception {
        if(args.getIsMetafile()) {
            // You can write your own method definition for converting Metafile to raster image using any third-party image converter.
            args.setImageStream(convertMetafileToRasterImage(args.getMetafileStream()));
        }else {
            InputStream inputStream = StreamSupport.toStream(args.getMetafileStream());
            // Use ByteArrayOutputStream to collect data into a byte array
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

            // Read data from the InputStream and write it to the ByteArrayOutputStream
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                byteArrayOutputStream.write(buffer, 0, bytesRead);
            }

            // Convert the ByteArrayOutputStream to a byte array
            byte[] tiffData = byteArrayOutputStream.toByteArray();
            // Read TIFF image from byte array
            ByteArrayInputStream tiffInputStream = new ByteArrayInputStream(tiffData);
            IIORegistry.getDefaultInstance().registerServiceProvider(new TIFFImageReaderSpi());

            // Create ImageReader and ImageWriter instances
            ImageReader tiffReader = ImageIO.getImageReadersByFormatName("TIFF").next();
            ImageWriter pngWriter = ImageIO.getImageWritersByFormatName("PNG").next();

            // Set up input and output streams
            tiffReader.setInput(ImageIO.createImageInputStream(tiffInputStream));
            ByteArrayOutputStream pngOutputStream = new ByteArrayOutputStream();
            pngWriter.setOutput(ImageIO.createImageOutputStream(pngOutputStream));

            // Read the TIFF image and write it as a PNG
            BufferedImage image = tiffReader.read(0);
            pngWriter.write(image);
            pngWriter.dispose();
            byte[] jpgData = pngOutputStream.toByteArray();
            InputStream jpgStream = new ByteArrayInputStream(jpgData);
            args.setImageStream(StreamSupport.toStream(jpgStream));
        }
    }

    // importFile helper method
    static FormatType getFormatType(String format)
    {
        switch (format)
        {
            case ".dotx":
            case ".docx":
            case ".docm":
            case ".dotm":
                return FormatType.Docx;
            case ".dot":
            case ".doc":
                return FormatType.Doc;
            case ".rtf":
                return FormatType.Rtf;
            case ".txt":
                return FormatType.Txt;
            case ".xml":
                return FormatType.WordML;
            case ".html":
                return FormatType.Html;
            default:
                return FormatType.Docx;
        }
    }

    // importFile() -> onMetafileImageParsed() -> helper method
    private static StreamSupport convertMetafileToRasterImage(StreamSupport ImageStream) throws Exception {
        //Here we are loading a default raster image as fallback.
        StreamSupport imgStream = getManifestResourceStream("ImageNotFound.jpg");
        return imgStream;
        //To do : Write your own logic for converting Metafile to raster image using any third-party image converter(Syncfusion doesn't provide any image converter).
    }

    // importFile() -> onMetafileImageParsed() -> convertMetafileToRasterImage() -> helper method
    private static StreamSupport getManifestResourceStream(String fileName) throws Exception {
        AssemblySupport assembly = AssemblySupport.getExecutingAssembly();
        return assembly.getManifestResourceStream("ImageNotFound.jpg");
    }

    // save() helper method
    static com.syncfusion.docio.FormatType getWFormatType(String format) throws Exception {
        if (format == null || format.trim().isEmpty())
            throw new Exception("EJ2 WordProcessor does not support this file format.");
        switch (format.toLowerCase()) {
            case ".dotx":
                return com.syncfusion.docio.FormatType.Dotx;
            case ".docm":
                return com.syncfusion.docio.FormatType.Docm;
            case ".dotm":
                return com.syncfusion.docio.FormatType.Dotm;
            case ".docx":
                return com.syncfusion.docio.FormatType.Docx;
            case ".rtf":
                return com.syncfusion.docio.FormatType.Rtf;
            case ".html":
                return com.syncfusion.docio.FormatType.Html;
            case ".txt":
                return com.syncfusion.docio.FormatType.Txt;
            case ".xml":
                return com.syncfusion.docio.FormatType.WordML;
            default:
                throw new Exception("EJ2 WordProcessor does not support this file format.");
        }
    }

}
