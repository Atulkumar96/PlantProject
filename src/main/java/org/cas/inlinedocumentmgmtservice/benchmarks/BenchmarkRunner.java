package org.cas.inlinedocumentmgmtservice.benchmarks;

public class BenchmarkRunner {
    public static void main(String[] args) throws Exception {
        // This call will discover and run all benchmarks (including DocumentServiceBenchmark)
        org.openjdk.jmh.Main.main(args);
    }
}
