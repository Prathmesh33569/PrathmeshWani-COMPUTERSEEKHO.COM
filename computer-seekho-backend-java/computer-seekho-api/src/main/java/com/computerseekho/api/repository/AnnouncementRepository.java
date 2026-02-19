package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Announcement;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.time.LocalDateTime;
import java.util.List;

public interface AnnouncementRepository
        extends JpaRepository<Announcement, Integer> {

    @Query("""
        SELECT a.announcementText
        FROM Announcement a
        WHERE a.isActive = true
          AND :now BETWEEN a.validFrom AND a.validTo
        ORDER BY a.createdAt DESC
    """)
    List<String> findActiveAnnouncementTexts(LocalDateTime now);
}
