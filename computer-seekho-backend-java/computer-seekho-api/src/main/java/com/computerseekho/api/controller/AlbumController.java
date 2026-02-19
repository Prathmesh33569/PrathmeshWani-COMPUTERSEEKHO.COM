package com.computerseekho.api.controller;

import com.computerseekho.api.dto.response.AlbumResponse;
import com.computerseekho.api.entity.Image;
import com.computerseekho.api.service.AlbumExcelService;
import com.computerseekho.api.service.AlbumImageUploadService;
import com.computerseekho.api.service.AlbumService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@CrossOrigin(origins = "*")
@RestController
@RequestMapping("/api/albums")
public class AlbumController {

    @Autowired
    private AlbumService albumService;


    @Autowired
    private AlbumExcelService albumExcelService;

    @Autowired
    private AlbumImageUploadService albumImageUploadService;

    /**
     * GET /api/albums/campus-life
     */
    @GetMapping("/campus-life")
    public ResponseEntity<List<AlbumResponse>> getCampusLifeAlbums() {
        return ResponseEntity.ok(albumService.getCampusLifeAlbums());
    }

    @PostMapping("/upload-excel")
    public ResponseEntity<?> uploadAlbumExcel(@RequestParam("file") MultipartFile file) {

        try {
            List<Image> images = albumExcelService.parseExcel(file);
            albumImageUploadService.saveAll(images);

            return ResponseEntity.ok(
                    "Excel uploaded successfully. Images saved: " + images.size()
            );

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body(e.getMessage());
        }
    }
}

