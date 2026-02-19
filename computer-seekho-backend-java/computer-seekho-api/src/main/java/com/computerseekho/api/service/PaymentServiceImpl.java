package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.PaymentCreateRequestDTO;
import com.computerseekho.api.dto.response.PaymentPdfDTO;
import com.computerseekho.api.dto.response.PaymentResponseDTO;
import com.computerseekho.api.entity.*;
import com.computerseekho.api.repository.*;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class PaymentServiceImpl implements PaymentService {

    private final PaymentRepository paymentRepository;
    private final ReceiptRepository receiptRepository;
    private final PaymentTypeRepository paymentTypeRepository;
    private final EnquiryRepository enquiryRepository;
    private final BatchRepository batchRepository;

    @Override
    @Transactional
    public PaymentResponseDTO createPayment(PaymentCreateRequestDTO dto) {

        // 1. Validate enquiry exists
        Enquiry enquiry = enquiryRepository.findById(dto.getEnquiryId())
                .orElseThrow(() -> new RuntimeException("Enquiry not found with id: " + dto.getEnquiryId()));

        // 2. Validate batch exists and get course fees
        Batch batch = batchRepository.findById(dto.getBatchId())
                .orElseThrow(() -> new RuntimeException("Batch not found with id: " + dto.getBatchId()));

        // 3. Validate payment type exists
        PaymentType paymentType = paymentTypeRepository.findById(dto.getPaymentTypeId())
                .orElseThrow(() -> new RuntimeException("Payment type not found with id: " + dto.getPaymentTypeId()));

        // 4. Validate amount matches course fees
//        BigDecimal courseFees = batch.getCourseFees();
//        if (courseFees == null) {
//            courseFees = batch.getCourse().getCourseFees(); // Fallback to course fees
//        }
//
//        if (!dto.getAmount().equals(courseFees)) {
//            throw new RuntimeException("Payment amount must match course fees: " + courseFees);
//        }

        // 4. Get course fees - USE THIS DIRECTLY (no validation needed)
        BigDecimal courseFees = batch.getCourseFees();
        if (courseFees == null) {
            courseFees = batch.getCourse().getCourseFees(); // Fallback to course fees
        }

        // 5. Create Payment record
        Payment payment = new Payment();
        payment.setPaymentTypeId(dto.getPaymentTypeId());
        payment.setPaymentDate(LocalDate.now());
        payment.setStudentId(null); // Will be updated after student registration
        payment.setCourseId(dto.getCourseId());
        payment.setBatchId(dto.getBatchId());
        payment.setAmount(dto.getAmount());
        payment.setEnquiryId(dto.getEnquiryId());

        Payment savedPayment = paymentRepository.save(payment);

        // 6. Create Receipt record
        Receipt receipt = new Receipt();
        receipt.setReceiptDate(LocalDate.now());
        receipt.setReceiptAmount(dto.getAmount());
        receipt.setPaymentId(savedPayment.getPaymentId());

        Receipt savedReceipt = receiptRepository.save(receipt);

        // 7. Map to response DTO
        return mapToResponseDTO(savedPayment, paymentType, savedReceipt);
    }

    @Override
    public Payment getPaymentById(Integer paymentId) {
        return paymentRepository.findById(paymentId)
                .orElseThrow(() -> new RuntimeException("Payment not found with id: " + paymentId));
    }

    @Override
    public List<PaymentResponseDTO> getAllPayments() {
        return paymentRepository.findAll().stream()
                .map(payment -> {
                    PaymentType paymentType = paymentTypeRepository.findById(payment.getPaymentTypeId()).orElse(null);
                    Receipt receipt = receiptRepository.findByPaymentId(payment.getPaymentId()).orElse(null);
                    return mapToResponseDTO(payment, paymentType, receipt);
                })
                .collect(Collectors.toList());
    }

    @Override
    public List<PaymentResponseDTO> getPaymentsByStudent(Integer studentId) {
        return paymentRepository.findByStudentId(studentId).stream()
                .map(payment -> {
                    PaymentType paymentType = paymentTypeRepository.findById(payment.getPaymentTypeId()).orElse(null);
                    Receipt receipt = receiptRepository.findByPaymentId(payment.getPaymentId()).orElse(null);
                    return mapToResponseDTO(payment, paymentType, receipt);
                })
                .collect(Collectors.toList());
    }

    private PaymentResponseDTO mapToResponseDTO(Payment payment, PaymentType paymentType, Receipt receipt) {
        PaymentResponseDTO dto = new PaymentResponseDTO();
        dto.setPaymentId(payment.getPaymentId());
        dto.setPaymentTypeId(payment.getPaymentTypeId());
        dto.setPaymentTypeDesc(paymentType != null ? paymentType.getPaymentTypeDesc() : null);
        dto.setPaymentDate(payment.getPaymentDate());
        dto.setAmount(payment.getAmount());
        dto.setStudentId(payment.getStudentId());
        dto.setCourseId(payment.getCourseId());
        dto.setBatchId(payment.getBatchId());
        dto.setEnquiryId(payment.getEnquiryId());
        dto.setReceiptId(receipt != null ? receipt.getReceiptId() : null);
        return dto;
    }

    @Override
    public PaymentPdfDTO getReceiptPdfData(Integer receiptId) {

        List<Object[]> rows = paymentRepository.getReceiptPdfData(receiptId);

        if (rows == null || rows.isEmpty()) {
            throw new RuntimeException("No receipt found for id: " + receiptId);
        }

        Object[] row = rows.get(0);

        PaymentPdfDTO dto = new PaymentPdfDTO();

        dto.setStudentName((String) row[0]);
        dto.setStudentMobile(row[1] != null ? row[1].toString() : null);
        dto.setStudentAddress((String) row[2]);
        dto.setStudentEmail((String) row[3]);
        dto.setCourseName((String) row[4]);
        dto.setAmount(((Number) row[5]).doubleValue());
        dto.setPaymentDate(((java.sql.Date) row[6]).toLocalDate());
        dto.setPaymentType((String) row[7]);
        dto.setReceiptAmount(((Number) row[8]).doubleValue());
        dto.setReceiptDate(((java.sql.Date) row[9]).toLocalDate());

        return dto;
    }



}