package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.PaymentCreateRequestDTO;
import com.computerseekho.api.dto.response.PaymentResponseDTO;
import com.computerseekho.api.service.PaymentService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/payments")
@RequiredArgsConstructor
@CrossOrigin(origins = "*")
public class PaymentController {

    private final PaymentService paymentService;

    /**
     * Create payment (before student registration)
     * POST /api/payments/create
     */
    @PostMapping("/create")
    public ResponseEntity<?> createPayment(@RequestBody PaymentCreateRequestDTO dto) {
        try {
            PaymentResponseDTO payment = paymentService.createPayment(dto);
            return ResponseEntity.status(HttpStatus.CREATED).body(payment);
        } catch (RuntimeException e) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(e.getMessage());
        }
    }

    /**
     * Get all payments
     * GET /api/payments
     */
    @GetMapping
    public ResponseEntity<List<PaymentResponseDTO>> getAllPayments() {
        List<PaymentResponseDTO> payments = paymentService.getAllPayments();
        return ResponseEntity.ok(payments);
    }

    /**
     * Get payments by student
     * GET /api/payments/student/{studentId}
     */
    @GetMapping("/student/{studentId}")
    public ResponseEntity<List<PaymentResponseDTO>> getPaymentsByStudent(@PathVariable Integer studentId) {
        List<PaymentResponseDTO> payments = paymentService.getPaymentsByStudent(studentId);
        return ResponseEntity.ok(payments);
    }
}