package com.computerseekho.api.dto.response;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class ReceiptResponseDTO {

    private Integer receiptId;
    private LocalDate receiptDate;
    private BigDecimal receiptAmount;
    private Integer paymentId;
}