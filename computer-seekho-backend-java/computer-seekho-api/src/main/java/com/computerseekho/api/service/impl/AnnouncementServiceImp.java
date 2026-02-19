package com.computerseekho.api.service.impl;

import com.computerseekho.api.entity.Announcement;
import com.computerseekho.api.repository.AnnouncementRepository;
import com.computerseekho.api.service.AnnouncementService;
import org.springframework.stereotype.Service;

import java.time.LocalDateTime;
import java.util.List;

@Service
public class AnnouncementServiceImp implements AnnouncementService
{
    private final AnnouncementRepository announcementRepository;

    public AnnouncementServiceImp(AnnouncementRepository announcementRepository) {
        this.announcementRepository = announcementRepository;
    }


    @Override
    public List<String> getActiveAnnouncementTexts() {
        return announcementRepository.findActiveAnnouncementTexts(LocalDateTime.now());

    }
    @Override
    public List<Announcement> getAllAnnouncements() {
        return announcementRepository.findAll();
    }

    @Override
    public Announcement getAnnouncementById(Integer id) {
        return announcementRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Announcement not found with id: " + id));
    }

    @Override
    public Announcement createAnnouncement(Announcement announcement) {
        return announcementRepository.save(announcement);
    }

    @Override
    public Announcement updateAnnouncement(Integer id, Announcement announcement) {
        Announcement existing = getAnnouncementById(id);

        existing.setAnnouncementText(announcement.getAnnouncementText());
        existing.setValidFrom(announcement.getValidFrom());
        existing.setValidTo(announcement.getValidTo());
        existing.setIsActive(announcement.getIsActive());

        return announcementRepository.save(existing);
    }

    @Override
    public void deleteAnnouncement(Integer id) {
        announcementRepository.deleteById(id);
    }
}
