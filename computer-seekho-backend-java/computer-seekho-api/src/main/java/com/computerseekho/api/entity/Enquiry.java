package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.UpdateTimestamp;

import java.time.LocalDate;
import java.time.LocalDateTime;

@Entity
@Table(name = "enquiry_master")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class Enquiry {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "enquiry_id")
    private Integer enquiryId;

    @Column(name = "enquirer_name", nullable = false, length = 100)
    private String enquirerName;

    @Column(name = "student_name", length = 100)
    private String studentName;

    @Column(name = "enquirer_address", length = 200)
    private String enquirerAddress;

    @Column(name = "enquirer_mobile", nullable = false)
    private Long enquirerMobile;

    @Column(name = "enquirer_alternate_mobile")
    private Long enquirerAlternateMobile;

    @Column(name = "enquirer_email_id", length = 100)
    private String enquirerEmailId;

    @Column(name = "enquiry_date")
    private LocalDate enquiryDate = LocalDate.now();

    @Column(name = "enquirer_query", columnDefinition = "TEXT")
    private String enquirerQuery;

    /* ===============================
       Foreign Key: Closure Reason
       =============================== */
    @ManyToOne
    @JoinColumn(name = "closure_reason_id")
    private ClosureReason closureReason;

    @Column(name = "closure_reason", length = 255)
    private String closureReasonText;

    @Column(name = "enquiry_processed_flag")
    private Boolean enquiryProcessedFlag = false;

    /* ===============================
       Foreign Key: Course
       =============================== */
    @ManyToOne
    @JoinColumn(name = "course_id")
    private Course course;

    /* ===============================
       Foreign Key: Staff
       =============================== */
    @ManyToOne
    @JoinColumn(name = "staff_id")
    private Staff staff;

    @Column(name = "inquiry_counter")
    private Integer inquiryCounter = 0;

    @Column(name = "followup_date")
    private LocalDate followupDate;

    @Column(name = "enquiry_source", length = 20)
    private String enquirySource; // Telephonic / Walk-in / Email / Online / Fax

    @Column(name = "special_instructions", columnDefinition = "TEXT")
    private String specialInstructions;

    @Column(name = "is_closed")
    private Boolean isClosed = false;

    @CreationTimestamp
    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    @UpdateTimestamp
    @Column(name = "updated_at")
    private LocalDateTime updatedAt;
}
