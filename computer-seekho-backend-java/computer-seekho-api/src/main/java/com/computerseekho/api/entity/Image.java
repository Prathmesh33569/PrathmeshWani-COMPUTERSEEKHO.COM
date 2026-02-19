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
@Table(name = "image_master")
public class Image {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "image_id")
    private Integer imageId;

    @Column(name = "image_path", nullable = false, length = 255)
    private String imagePath;

    // FK → album_master
    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "album_id", nullable = false)
    private Album album;

    @Column(name = "is_album_cover")
    private Boolean isAlbumCover = false;

    @Column(name = "image_is_active")
    private Boolean imageIsActive = true;

    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    // ---------- Audit ----------

    @PrePersist
    protected void onCreate() {
        createdAt = LocalDateTime.now();
    }
}