package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.*;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;


@Data
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Entity
@Table(name = "course_master")

public class Course {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "course_id")
    private Integer courseId;

    @Column(name = "course_name", nullable = false, length = 100)
    private String courseName;

    @Column(name = "course_description", columnDefinition = "TEXT")
    private String courseDescription;

    @Column(name = "course_duration")
    private Integer courseDuration; // duration in months

    @Column(name = "course_fees", nullable = false, precision = 10, scale = 2)
    private BigDecimal courseFees;

    @Column(name = "course_fees_from")
    private LocalDate courseFeesFrom;

    @Column(name = "course_fees_to")
    private LocalDate courseFeesTo;

    @Column(name = "course_syllabus", columnDefinition = "TEXT")
    private String courseSyllabus;

    @Column(name = "age_grp_type", length = 20)
    private String ageGrpType; // Adult / All Ages / Kids

    @Column(name = "course_is_active")
    private Boolean courseIsActive = true;

    @Column(name = "cover_photo", length = 255)
    private String coverPhoto;

    @Column(name = "video_id")
    private Integer videoId; // FK (video_master)

    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    @Column(name = "updated_at")
    private LocalDateTime updatedAt;

    /* =======================
       JPA Lifecycle Callbacks
       ======================= */

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
