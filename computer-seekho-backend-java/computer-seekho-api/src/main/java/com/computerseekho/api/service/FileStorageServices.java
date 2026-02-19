package com.computerseekho.api.service;

import com.computerseekho.api.config.FileStorageProperties;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.UUID;

@Service
public class FileStorageServices {

    private final Path fileStorageLocation;

    @Autowired
    public FileStorageServices(FileStorageProperties fileStorageProperties) {
        // Use an external directory for file storage (NOT inside src/main/resources)
        // This ensures files are immediately accessible without restart
        this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir())
                .toAbsolutePath().normalize();

        try {
            Files.createDirectories(this.fileStorageLocation);
            System.out.println("✅ File storage directory created/verified at: " + this.fileStorageLocation);
        } catch (Exception ex) {
            throw new RuntimeException("Could not create the directory where the uploaded files will be stored.", ex);
        }
    }

    /**
     * Store staff image with unique filename
     */
    public String storeStaffImage(MultipartFile file, String staffUsername) {
        // Get original filename
        String originalFileName = StringUtils.cleanPath(file.getOriginalFilename());

        try {
            // Check if the file contains invalid characters
            if (originalFileName.contains("..")) {
                throw new RuntimeException("Invalid file path: " + originalFileName);
            }

            // Generate unique filename to avoid conflicts
            String fileExtension = "";
            if (originalFileName.contains(".")) {
                fileExtension = originalFileName.substring(originalFileName.lastIndexOf("."));
            }

            // Create unique filename: staff_username_timestamp.ext
            String uniqueFileName = "staff_" + staffUsername + "_" + System.currentTimeMillis() + fileExtension;

            // Copy file to the target location (Replacing existing file with the same name)
            Path targetLocation = this.fileStorageLocation.resolve(uniqueFileName);
            Files.copy(file.getInputStream(), targetLocation, StandardCopyOption.REPLACE_EXISTING);

            System.out.println("✅ Image stored successfully: " + uniqueFileName);

            // Return the relative URL path (not full filesystem path)
            return "/assets/staff/" + uniqueFileName;

        } catch (IOException ex) {
            throw new RuntimeException("Could not store file " + originalFileName + ". Please try again!", ex);
        }
    }

    /**
     * Validate file size (max 5MB)
     */
    public boolean isValidFileSize(MultipartFile file) {
        long maxFileSize = 5 * 1024 * 1024; // 5MB in bytes
        return file.getSize() <= maxFileSize;
    }

    /**
     * Get file storage location
     */
    public Path getFileStorageLocation() {
        return fileStorageLocation;
    }
}