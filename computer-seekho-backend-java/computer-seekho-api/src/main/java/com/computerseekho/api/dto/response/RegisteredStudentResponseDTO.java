package com.computerseekho.api.dto.response;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

@Data
public class RegisteredStudentResponseDTO {

    // Student Info
    private Integer studentId;
    private String studentName;
    private String studentMobile;
    private String studentGender;
    private LocalDate studentDob;
    private String studentQualification;
    private String studentAddress;
    private String photoUrl;

    // Course Info
    private Integer courseId;
    private String courseName;

    // Batch Info
    private Integer batchId;
    private String batchName;

    // Payment Info
    private Integer paymentId;
    private BigDecimal amount;
    private LocalDate paymentDate;
    private Integer paymentTypeId;
    private String paymentTypeDesc;

    // Receipt Info
    private Integer receiptId;
    private LocalDate receiptDate;
    private BigDecimal receiptAmount;

    // Additional Info
    private Integer enquiryId;
    private LocalDateTime registeredAt; // Student creation time
}