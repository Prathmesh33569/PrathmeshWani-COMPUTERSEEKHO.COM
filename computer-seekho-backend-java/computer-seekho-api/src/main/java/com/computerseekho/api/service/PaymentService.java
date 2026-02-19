package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.PaymentCreateRequestDTO;
import com.computerseekho.api.dto.response.PaymentPdfDTO;
import com.computerseekho.api.dto.response.PaymentResponseDTO;
import com.computerseekho.api.entity.Payment;

import java.util.List;

public interface PaymentService {

    PaymentResponseDTO createPayment(PaymentCreateRequestDTO dto);

    Payment getPaymentById(Integer paymentId);

    List<PaymentResponseDTO> getAllPayments();

    List<PaymentResponseDTO> getPaymentsByStudent(Integer studentId);
    public PaymentPdfDTO getReceiptPdfData(Integer receiptId);
}