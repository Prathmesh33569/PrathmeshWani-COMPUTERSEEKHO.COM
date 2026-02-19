package com.computerseekho.api.dto.response;

import lombok.Data;

import java.time.LocalDate;

@Data
public class PaymentPdfDTO {

    private String studentName;
    private String studentMobile;
    private String studentAddress;
    private String studentEmail;
    private String courseName;
    private Double amount;
    private LocalDate paymentDate;
    private String paymentType;
    private Double receiptAmount;
    private LocalDate receiptDate;
}
