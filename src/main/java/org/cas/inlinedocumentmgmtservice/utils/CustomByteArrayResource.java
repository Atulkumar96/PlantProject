package org.cas.inlinedocumentmgmtservice.utils;

import org.springframework.core.io.ByteArrayResource;

/**
 * CustomByteArrayResource class
 * Spring’s ByteArrayResource does not include a filename by default when sending multipart files.
 * This class extends ByteArrayResource and overrides the getFilename method to include the filename.
 * This class is used to include the filename when sending multipart files.
 */

public class CustomByteArrayResource extends ByteArrayResource {
    private final String fileName;

    public CustomByteArrayResource(byte[] byteArray, String fileName) {
        super(byteArray);
        this.fileName = fileName;
    }

    @Override
    public String getFilename() {
        return this.fileName;
    }
}
