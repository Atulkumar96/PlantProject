package org.cas.inlinedocumentmgmtservice.controllers;


import com.syncfusion.docio.WordDocument;
import com.syncfusion.javahelper.system.io.StreamSupport;
import com.syncfusion.javahelper.system.reflection.AssemblySupport;
import com.twelvemonkeys.imageio.plugins.tiff.TIFFImageReaderSpi;
import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.dtos.ResponseDto;
import org.cas.inlinedocumentmgmtservice.dtos.SaveDto;
import org.cas.inlinedocumentmgmtservice.services.DocumentService;
import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import org.springframework.web.bind.annotation.*;

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


import org.springframework.web.bind.annotation.*;

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

import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;
import com.syncfusion.ej2.wordprocessor.WordProcessorHelper;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
//import com.syncfusion.javahelper.system.io.StreamSupport;
//import com.syncfusion.javahelper.system.reflection.AssemblySupport;
import com.twelvemonkeys.imageio.plugins.tiff.TIFFImageReaderSpi;
//import com.google.gson.Gson;
//import com.google.gson.JsonArray;
//import com.google.gson.JsonObject;
import com.syncfusion.ej2.spellchecker.DictionaryData;
import com.syncfusion.ej2.spellchecker.SpellChecker;
//import com.syncfusion.docio.WordDocument;
import com.syncfusion.ej2.wordprocessor.FormatType;
//import
import com.syncfusion.ej2.wordprocessor.MetafileImageParsedEventArgs;
import com.syncfusion.ej2.wordprocessor.MetafileImageParsedEventHandler;


@RestController
@RequestMapping("/api/inlineDocumentService")
public class DocumentController {
    private final DocumentService documentService;
    private final DocumentServiceImpl documentServiceImpl;

    public DocumentController(DocumentService documentService, DocumentServiceImpl documentServiceImpl) {
        this.documentService = documentService;
        this.documentServiceImpl = documentServiceImpl;
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

    @CrossOrigin(origins = "*", allowedHeaders = "*")
    @PostMapping("Import")
    public String importFile(@RequestParam("files") MultipartFile file) throws Exception {
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
            try {
                sfdtContent = WordProcessorHelper.load(file.getInputStream(), getFormatType(format));
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

    @CrossOrigin(origins = "*", allowedHeaders = "*")
    @PostMapping("Save")
    public void save(@RequestBody SaveDto data) throws Exception {
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

    // Convert String sfdt to docx
    private void convertSfdtToDocx(String sfdtContent, String filePath) throws Exception {
        WordDocument document = WordProcessorHelper.save(sfdtContent);
        FileOutputStream fileStream = new FileOutputStream(filePath);
        document.save(fileStream, com.syncfusion.docio.FormatType.Docx);
        fileStream.close();
        document.close();
    }
}
