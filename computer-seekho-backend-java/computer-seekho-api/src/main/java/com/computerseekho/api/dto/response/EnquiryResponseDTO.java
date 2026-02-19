package com.computerseekho.api.dto.response;

import lombok.Data;

import java.time.LocalDate;

@Data
public class EnquiryResponseDTO {

    private Integer enquiryId;
    private String enquirerName;
    private Long enquirerMobile;
    private String courseName;
    private LocalDate followupDate;
    private Boolean isClosed;
}
