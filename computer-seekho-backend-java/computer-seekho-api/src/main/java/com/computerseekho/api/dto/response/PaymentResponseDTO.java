package com.computerseekho.api.dto.response;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class PaymentResponseDTO {

    private Integer paymentId;
    private Integer paymentTypeId;
    private String paymentTypeDesc;
    private LocalDate paymentDate;
    private BigDecimal amount;
    private Integer studentId;
    private Integer courseId;
    private Integer batchId;
    private Integer enquiryId;
    private Integer receiptId;      // Linked receipt ID
}