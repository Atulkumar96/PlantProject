package org.cas.inlinedocumentmgmtservice.benchmarks;

import org.cas.inlinedocumentmgmtservice.dtos.PlantDto;
import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;
import org.openjdk.jmh.annotations.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.concurrent.TimeUnit;

/**
 * This class contains JMH benchmarks for the DocumentService class.
 * It measures the performance of various methods in the DocumentService class.
 * The benchmarks are executed repeatedly to measure the average time taken for each method.
 * The setup method initializes the service and the test document path.
 * The benchmark methods measure the performance of the methods in the DocumentService class.
 * The SimpleMultipartFile class is a simple implementation of the MultipartFile interface for benchmarking purposes.
 * The main method in the BenchmarkRunner class is used to run the benchmarks.
 *
 * TODO: Ensure to set the @BenchmarkMode(Mode.SingleShotTime) for each benchmark method.
 */

@State(Scope.Benchmark)
public class DocumentServiceBenchmark {
    private DocumentServiceImpl documentServiceImpl;
    private String documentPath;

    // Number of comments to add to the document
    @Param({"10", "50", "100"})
    public int numComments;

    @Setup(Level.Trial)
    public void setup() throws IOException {
        try{
            // Copy the resource DOCX from your classpath to a temporary file.
            InputStream resourceStream = this.getClass().getResourceAsStream("/RSAW BAL-001-2_2016_v1.docx");
            if (resourceStream == null) {
                throw new RuntimeException("Resource file not found.");
            }
            File tempFile = File.createTempFile("testDoc", ".docx");
            tempFile.deleteOnExit();

            Files.copy(resourceStream, tempFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            this.documentPath = tempFile.getAbsolutePath();

            // Initialize the service. Pass null for ResourceLoader if not used.
            documentServiceImpl = new DocumentServiceImpl(null);
        }
        catch (Exception e){
            throw new RuntimeException("Failed to set up benchmark", e);
        }
    }

    // Benchmark methods that will be executed repeatedly.
    // It measures the average time to extract review comments from the document.

    //==========================================================================
    // 1. Benchmark addComments(String, int)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkAddComments() throws Exception {
        // Create a temporary copy of the original document for a clean slate.
        File tempDoc = File.createTempFile("testDocAddComments", ".docx");
        tempDoc.deleteOnExit();
        Files.copy(new File(documentPath).toPath(), tempDoc.toPath(), StandardCopyOption.REPLACE_EXISTING);

        // Call the addComments method with the current parameter value.
        documentServiceImpl.addComments(tempDoc.getAbsolutePath(), numComments);
    }


    //==========================================================================
    // 2. Benchmark extractReviewComments(String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public String benchmarkExtractReviewComments() throws Exception {
        return documentServiceImpl.extractReviewComments(documentPath);
    }

    //==========================================================================
    // 3. Benchmark extractReviewComments(MultipartFile)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public String benchmarkExtractReviewCommentsMultipart() throws Exception {
        File file = new File(documentPath);
        SimpleMultipartFile multipartFile = new SimpleMultipartFile(file.getName(), Files.readAllBytes(file.toPath()));
        return documentServiceImpl.extractReviewComments(multipartFile);
    }

    //==========================================================================
    // 4. Benchmark mailMerge(PlantDto)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkMailMerge() throws Exception {
        // Create a simple PlantDto instance.
        //PlantDto plantDto = new PlantDto();
        String plantName = "Electric Plant";
        documentServiceImpl.mailMerge(plantName);
    }

    //==========================================================================
    // 5. Benchmark appendSignature(String, String, String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkAppendSignature() throws Exception {
        File outputFile = File.createTempFile("appendSignatureOutput", ".docx");
        outputFile.deleteOnExit();

        documentServiceImpl.appendSignature(documentPath, outputFile.getAbsolutePath(), "ApproverName");
    }

    //==========================================================================
    // 6. Benchmark appendSignature(MultipartFile, String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public String benchmarkAppendSignatureMultipart() throws Exception {
        File file = new File(documentPath);
        SimpleMultipartFile multipartFile = new SimpleMultipartFile(file.getName(), Files.readAllBytes(file.toPath()));
        return documentServiceImpl.appendSignature(multipartFile, "ApproverName");
    }


    //==========================================================================
    // 7. Benchmark insertLink(String, String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkInsertLink() throws Exception {
        File outputFile = File.createTempFile("insertLinkOutput", ".docx");
        outputFile.deleteOnExit();
        documentServiceImpl.insertLink(documentPath, outputFile.getAbsolutePath());
    }

    //==========================================================================
    // 8. Benchmark insertOle(String, String, String, String, String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkInsertOle() throws Exception {
        File outputFile = File.createTempFile("insertOleOutput", ".docx");
        outputFile.deleteOnExit();

        // Check test files for PDF, Excel, and DOC are in resources.
        InputStream pdfStream = this.getClass().getResourceAsStream("/testPdf.pdf");
        InputStream excelStream = this.getClass().getResourceAsStream("/testExcel.xlsx");
        InputStream docStream = this.getClass().getResourceAsStream("/testDoc.docx");
        if (pdfStream == null || excelStream == null || docStream == null) {
            throw new RuntimeException("One or more resource files for OLE not found.");
        }
        File pdfFile = File.createTempFile("testPdf", ".pdf");
        pdfFile.deleteOnExit();
        Files.copy(pdfStream, pdfFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        File excelFile = File.createTempFile("testExcel", ".xlsx");
        excelFile.deleteOnExit();
        Files.copy(excelStream, excelFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        File docFile = File.createTempFile("testDocOle", ".docx");
        docFile.deleteOnExit();
        Files.copy(docStream, docFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        documentServiceImpl.insertOle(documentPath, outputFile.getAbsolutePath(),
                pdfFile.getAbsolutePath(), excelFile.getAbsolutePath(), docFile.getAbsolutePath());
    }


    //==========================================================================
    // 9. Benchmark insertImage(String, String, String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkInsertImageWithPaths() throws Exception {
        File outputFile = File.createTempFile("insertImageOutput", ".docx");
        outputFile.deleteOnExit();
        // Assume a test image is available in resources.
        InputStream imageStream = this.getClass().getResourceAsStream("/testImage.jpg");
        if (imageStream == null) {
            throw new RuntimeException("Image resource not found.");
        }
        File imageFile = File.createTempFile("testImage", ".jpg");
        imageFile.deleteOnExit();
        Files.copy(imageStream, imageFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        documentServiceImpl.insertImage(documentPath, imageFile.getAbsolutePath(), outputFile.getAbsolutePath());
    }

    //==========================================================================
    // 10. Benchmark insertImage(MultipartFile, MultipartFile)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public String benchmarkInsertImageMultipart() throws Exception {
        File file = new File(documentPath);
        SimpleMultipartFile docMultipart = new SimpleMultipartFile(file.getName(), Files.readAllBytes(file.toPath()));

        InputStream imageStream = this.getClass().getResourceAsStream("/testImage.jpg");
        if (imageStream == null) {
            throw new RuntimeException("Image resource not found.");
        }
        byte[] imageBytes = imageStream.readAllBytes();
        SimpleMultipartFile imageMultipart = new SimpleMultipartFile("testImage.jpg", imageBytes);

        return documentServiceImpl.insertImage(docMultipart, imageMultipart);
    }


    //==========================================================================
    // 11. Benchmark protectDocument(String, String)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkProtectDocumentWithPaths() throws Exception {
        File outputFile = File.createTempFile("protectOutput", ".docx");
        outputFile.deleteOnExit();
        documentServiceImpl.protectDocument(documentPath, outputFile.getAbsolutePath());
    }

    //==========================================================================
    // 12. Benchmark protectDocument(MultipartFile)
    //==========================================================================
    @Warmup(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Measurement(iterations = 3, time = 3, timeUnit = TimeUnit.SECONDS)
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkProtectDocumentMultipart() throws Exception {
        File file = new File(documentPath);
        SimpleMultipartFile multipartFile = new SimpleMultipartFile(file.getName(), Files.readAllBytes(file.toPath()));
        documentServiceImpl.protectDocument(multipartFile);
    }


    //==========================================================================
    // A simple MultipartFile implementation for benchmarking purposes.
    //==========================================================================
    public static class SimpleMultipartFile implements MultipartFile {
        private final String name;
        private final byte[] content;

        public SimpleMultipartFile(String name, byte[] content) {
            this.name = name;
            this.content = content;
        }

        @Override
        public String getName() {
            return name;
        }

        @Override
        public String getOriginalFilename() {
            return name;
        }

        @Override
        public String getContentType() {
            return "application/octet-stream";
        }

        @Override
        public boolean isEmpty() {
            return content == null || content.length == 0;
        }

        @Override
        public long getSize() {
            return content.length;
        }

        @Override
        public byte[] getBytes() throws IOException {
            return content;
        }

        @Override
        public InputStream getInputStream() throws IOException {
            return new ByteArrayInputStream(content);
        }

        @Override
        public void transferTo(File dest) throws IOException, IllegalStateException {
            Files.write(dest.toPath(), content);
        }
    }

}

/**
 * @State(Scope.Benchmark):
 * The state (here, our service instance and the test document path) is shared across all benchmark iterations.
 *
 * @Setup(Level.Trial):
 * The setup method runs once per trial. It copies a DOCX file from your resources into a temporary file and initializes your service.
 *
 * @Benchmark:
 * This annotation marks the method that JMH will repeatedly call to measure its performance.
 *
 * BenchmarkMode and OutputTimeUnit:
 * In this example, we measure the average time taken (in milliseconds) for each call.
     * Mode.Throughput – Measures the number of operations per second.
     * Mode.AverageTime – Measures the average execution time per operation. This is ideal when you want to know the typical latency of an operation.
     * Mode.SampleTime – Measures execution time distribution.
     * Mode.SingleShotTime – Measures the execution time of a single invocation.
     * Mode.All – Runs all modes.
 */

/**
 * To run the benchmark, you can use the BenchmarkRunner class. This class is a simple wrapper around JMH that allows you to run your benchmarks from the command line.
 */
