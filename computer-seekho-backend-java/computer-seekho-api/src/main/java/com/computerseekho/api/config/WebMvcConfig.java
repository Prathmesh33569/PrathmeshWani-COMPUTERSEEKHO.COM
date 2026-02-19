package com.computerseekho.api.config;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class WebMvcConfig implements WebMvcConfigurer {

    @Autowired
    private FileStorageProperties fileStorageProperties;

    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {
        // Get the absolute path for file storage
        String uploadDir = fileStorageProperties.getUploadDir();

        // Serve static images from external directory (NOT from classpath)
        // This allows files to be accessible immediately after upload without restart
        registry.addResourceHandler("/assets/staff/**")
                .addResourceLocations("file:" + uploadDir + "/")
                .setCachePeriod(0)  // Disable caching for immediate access to new files
                .resourceChain(false);  // Disable resource chain for dynamic content

        System.out.println("✅ Static resource handler configured for /assets/staff/");
        System.out.println("✅ Serving files from: file:" + uploadDir + "/");
        System.out.println("✅ Cache disabled for immediate file access");
    }
}