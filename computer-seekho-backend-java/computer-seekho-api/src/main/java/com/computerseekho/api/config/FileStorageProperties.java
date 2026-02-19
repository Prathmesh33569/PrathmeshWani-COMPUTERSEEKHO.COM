package com.computerseekho.api.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@ConfigurationProperties(prefix = "file.storage")
public class FileStorageProperties {

    // IMPORTANT: Use an external directory (NOT inside src/main/resources)
    // This ensures uploaded files are immediately accessible without restart
    // Example Windows: C:/uploads/staff
    // Example Linux: /var/uploads/staff
    // Example Development: ./uploads/staff (relative to project root)
    private String uploadDir = "./uploads/staff";

    public String getUploadDir() {
        return uploadDir;
    }

    public void setUploadDir(String uploadDir) {
        this.uploadDir = uploadDir;
    }
}