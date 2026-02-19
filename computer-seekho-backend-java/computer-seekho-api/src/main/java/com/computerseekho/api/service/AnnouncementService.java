package com.computerseekho.api.service;

import com.computerseekho.api.entity.Announcement;
import org.springframework.stereotype.Service;

import java.util.List;

public interface AnnouncementService {

    List<String> getActiveAnnouncementTexts();
    List<Announcement> getAllAnnouncements();
    Announcement getAnnouncementById(Integer id);
    Announcement createAnnouncement(Announcement announcement);
    Announcement updateAnnouncement(Integer id, Announcement announcement);
    void deleteAnnouncement(Integer id);
}
