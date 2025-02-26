package org.cas.inlinedocumentmgmtservice.benchmarks;

import org.cas.inlinedocumentmgmtservice.services.DocumentServiceImpl;
import org.openjdk.jmh.annotations.*;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.concurrent.TimeUnit;

@State(Scope.Benchmark)
public class DocumentServiceBenchmark {
    private DocumentServiceImpl documentServiceImpl;
    private String documentPath;

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

            // Initialize the service. Pass null for ResourceLoader if it’s not used in this method.
            documentServiceImpl = new DocumentServiceImpl(null);
        }
        catch (Exception e){
            throw new RuntimeException("Failed to set up benchmark", e);
        }
    }

    // Benchmark methods that will be executed repeatedly.
    // It measures the average time to extract review comments from the document.
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public String benchmarkExtractReviewComments() throws Exception {
        return documentServiceImpl.extractReviewComments(documentPath);
    }

    // It measures the average time to insert a link into the document.
    @Benchmark
    @BenchmarkMode(Mode.AverageTime)
    @OutputTimeUnit(TimeUnit.MILLISECONDS)
    public void benchmarkInsertLink() throws Exception {
        // Create a temporary file to store the output.
        File.createTempFile("insertLinkOutput", ".docx").deleteOnExit();

        // Call the method under test.
        documentServiceImpl.insertLink(documentPath, documentPath);
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
