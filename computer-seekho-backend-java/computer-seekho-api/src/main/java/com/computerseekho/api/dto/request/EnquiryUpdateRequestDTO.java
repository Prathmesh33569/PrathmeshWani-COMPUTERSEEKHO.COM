package com.computerseekho.api.dto.request;

import lombok.Data;

import java.time.LocalDate;

@Data
public class EnquiryUpdateRequestDTO {

    private String enquirerName;
    private String studentName;
    private String enquirerAddress;
    private Long enquirerMobile;
    private Long enquirerAlternateMobile;
    private String enquirerEmailId;

    private String enquirerQuery;
    private String enquirySource;   // Telephonic / Walk-in / Email / Online / Fax

    private Integer courseId;        // Course can be changed

    private LocalDate followupDate;  // Can update follow-up date
    private String specialInstructions;
}