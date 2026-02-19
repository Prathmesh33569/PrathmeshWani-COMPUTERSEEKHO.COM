package com.computerseekho.api.dto.request;

import lombok.Data;

import java.time.LocalDate;

@Data
public class EnquiryCreateRequestDTO {

    private String enquirerName;
    private String studentName;
    private String enquirerAddress;
    private Long enquirerMobile;
    private Long enquirerAlternateMobile;
    private String enquirerEmailId;

    private String enquirerQuery;
    private String enquirySource;   // Telephonic / Walk-in / Email / Online / Fax

    private Integer courseId;        // Course A (initial enquiry)
    private Integer staffId;         // Logged-in staff

    private LocalDate followupDate;  // Optional (default +3 days)
    private String specialInstructions;
}
