package com.computerseekho.api.dto.response;

import lombok.Data;

import java.time.LocalDate;

@Data
public class StudentResponseDTO {

    private Integer studentId;
    private String studentName;
    private String studentMobile;
    private String studentGender;
    private LocalDate studentDob;
    private String studentQualification;
    private String studentAddress;
    private String photoUrl;

    private Integer batchId;
    private String batchName;       // From Batch entity

    private Integer courseId;
    private String courseName;      // From Course entity

    private String studentUsername;

    private Integer enquiryId;      // Reference to original enquiry
}