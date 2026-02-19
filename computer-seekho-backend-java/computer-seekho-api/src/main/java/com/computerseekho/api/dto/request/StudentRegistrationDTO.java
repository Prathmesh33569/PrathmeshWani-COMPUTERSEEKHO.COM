package com.computerseekho.api.dto.request;

import lombok.Data;

import java.time.LocalDate;

@Data
public class StudentRegistrationDTO {

    // From Enquiry (pre-filled)
    private Integer enquiryId;          // Required - to mark enquiry as processed

    // Student Basic Details
    private String studentName;         // Pre-filled from enquiry
    private Long studentMobile;         // Pre-filled from enquiry
    private String studentAddress;      // Pre-filled from enquiry if available

    // Additional Student Details (user enters)
    private String studentGender;       // Male/Female/Other
    private LocalDate studentDob;       // Date of Birth
    private String studentQualification; // 10th/12th/Graduate/Post-Graduate
    private String photoUrl;            // Optional - photo upload

    // Course & Batch
    private Integer courseId;           // Pre-filled from enquiry
    private Integer batchId;            // User selects batch

    // Login Credentials
    private String studentUsername;     // User enters or auto-generated
    private String studentPassword;     // User enters

    // Payment Details - NOW REQUIRED
    private Integer paymentId;          // Payment ID from payment creation step
}