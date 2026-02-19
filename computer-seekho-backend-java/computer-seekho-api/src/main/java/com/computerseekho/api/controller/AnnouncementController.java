package com.computerseekho.api.controller;


import com.computerseekho.api.entity.Announcement;
import com.computerseekho.api.service.AnnouncementService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/announcements")
@CrossOrigin(origins = "*")
public class AnnouncementController {

    private final AnnouncementService announcementService;

    public AnnouncementController(AnnouncementService announcementService) {
        this.announcementService = announcementService;
    }

    @GetMapping("/active")
    public List<String> getActiveAnnouncements() {
        return announcementService.getActiveAnnouncementTexts();
    }

    @GetMapping
    public List<Announcement> getAllAnnouncements() {
        return announcementService.getAllAnnouncements();
    }

    @GetMapping("/{id}")
    public ResponseEntity<Announcement> getAnnouncementById(@PathVariable Integer id) {
        return ResponseEntity.ok(announcementService.getAnnouncementById(id));
    }

    @PostMapping
    public ResponseEntity<Announcement> createAnnouncement(@RequestBody Announcement announcement) {
        return ResponseEntity.status(HttpStatus.CREATED)
                .body(announcementService.createAnnouncement(announcement));
    }

    @PutMapping("/{id}")
    public ResponseEntity<Announcement> updateAnnouncement(
            @PathVariable Integer id,
            @RequestBody Announcement announcement) {
        return ResponseEntity.ok(announcementService.updateAnnouncement(id, announcement));
    }

    @DeleteMapping("/{id}")
    public ResponseEntity<Void> deleteAnnouncement(@PathVariable Integer id) {
        announcementService.deleteAnnouncement(id);
        return ResponseEntity.noContent().build();
    }
}
