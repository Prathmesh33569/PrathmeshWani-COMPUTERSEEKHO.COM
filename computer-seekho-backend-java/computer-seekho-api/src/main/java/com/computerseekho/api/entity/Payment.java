package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Entity
@Table(name = "payment")
public class Payment {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "payment_id")
    private Integer paymentId;

    @Column(name = "payment_type_id", nullable = false)
    private Integer paymentTypeId; // FK to payment_type_master

    @Column(name = "payment_date", nullable = false)
    private LocalDate paymentDate;

    @Column(name = "student_id")
    private Integer studentId; // FK to student_master

    @Column(name = "course_id", nullable = false)
    private Integer courseId; // FK to course_master

    @Column(name = "batch_id", nullable = false)
    private Integer batchId; // FK to batch_master

    @Column(name = "amount", nullable = false, precision = 10, scale = 2)
    private BigDecimal amount;

    @Column(name = "enquiry_id")
    private Integer enquiryId; // FK to enquiry

    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    @Column(name = "updated_at")
    private LocalDateTime updatedAt;

    @PrePersist
    protected void onCreate() {
        createdAt = LocalDateTime.now();
        updatedAt = LocalDateTime.now();
        if (paymentDate == null) {
            paymentDate = LocalDate.now();
        }
    }

    @PreUpdate
    protected void onUpdate() {
        updatedAt = LocalDateTime.now();
    }
}