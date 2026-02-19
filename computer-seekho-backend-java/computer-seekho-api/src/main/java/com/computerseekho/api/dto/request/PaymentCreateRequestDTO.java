package com.computerseekho.api.dto.request;

import lombok.Data;

import java.math.BigDecimal;

@Data
public class PaymentCreateRequestDTO {

    private Integer enquiryId;      // Required
    private Integer courseId;       // Required
    private Integer batchId;        // Required
    private Integer paymentTypeId;  // Required (1=Cheque, 2=DD, 3=Bank Transfer, 4=Cash)
    private BigDecimal amount;      // Required (from course fees)
}