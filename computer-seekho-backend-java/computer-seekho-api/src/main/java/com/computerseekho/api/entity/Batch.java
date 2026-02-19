package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

@Getter
@Setter
@NoArgsConstructor
@Entity
@Table(name = "batch_master")
public class Batch {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "batch_id")
    private Integer batchId;

    @Column(name = "batch_name", nullable = false, length = 100)
    private String batchName;

    @Column(name = "batch_start_date")
    private LocalDate batchStartDate;

    @Column(name = "batch_end_date")
    private LocalDate batchEndDate;

    // FK → course_master
    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "course_id", nullable = false)
    private Course course;

    @Column(name = "presentation_date")
    private LocalDateTime presentationDate;

    @Column(name = "course_fees", precision = 10, scale = 2)
    private BigDecimal courseFees;

    @Column(name = "course_fees_from")
    private LocalDate courseFeesFrom;

    @Column(name = "course_fees_to")
    private LocalDate courseFeesTo;

    @Column(name = "batch_is_active")
    private Boolean batchIsActive = true;

    @Column(name = "final_presentation_date")
    private LocalDateTime finalPresentationDate;

    @Column(name = "batch_logo_url", length = 255)
    private String batchLogoUrl;

    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    @Column(name = "updated_at")
    private LocalDateTime updatedAt;

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