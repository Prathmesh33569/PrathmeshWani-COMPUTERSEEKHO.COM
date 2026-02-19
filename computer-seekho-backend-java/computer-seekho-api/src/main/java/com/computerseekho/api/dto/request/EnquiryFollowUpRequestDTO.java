package com.computerseekho.api.dto.request;

import lombok.Data;

import java.time.LocalDate;

@Data
public class EnquiryFollowUpRequestDTO {

    private Integer enquiryId;

    private String remarks;
    private String specialInstructions;

    private LocalDate nextFollowupDate;  // default today + 3
    private Boolean closeEnquiry;

    private Integer closureReasonId;     // optional
    private String closureReasonText;    // if "Other"
}
