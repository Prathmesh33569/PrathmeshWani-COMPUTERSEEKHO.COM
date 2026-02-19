package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.time.LocalDateTime;

@Getter
@Setter
@NoArgsConstructor
@Entity
@Table(name = "album_master")
public class Album {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "album_id")
    private Integer albumId;

    // PLACEMENT / CAMPUS_LIFE / RECRUITER
    @Enumerated(EnumType.STRING)
    @Column(name = "album_type", nullable = false, length = 30)
    private AlbumType albumType;

    // DAC / DBDA / NULL (for campus life)
    @Enumerated(EnumType.STRING)
    @Column(name = "program_code", length = 20)
    private ProgramCode programCode;

    @Column(name = "album_name", nullable = false, length = 150)
    private String albumName;

    @Column(name = "album_description", length = 255)
    private String albumDescription;

    // Batch / event start date
    @Column(name = "start_date")
    private LocalDateTime startDate;

    // Batch / event end date
    @Column(name = "end_date")
    private LocalDateTime endDate;

    // 1 = On, 0 = Off
    @Column(name = "album_is_active")
    private Boolean albumIsActive = true;

    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    @Column(name = "updated_at")
    private LocalDateTime updatedAt;

    // ---------- Audit ----------
    @PrePersist
    protected void onCreate() {
        createdAt = LocalDateTime.now();
        updatedAt = LocalDateTime.now();
    }

    @PreUpdate
    protected void onUpdate() {
        updatedAt = LocalDateTime.now();
    }
}