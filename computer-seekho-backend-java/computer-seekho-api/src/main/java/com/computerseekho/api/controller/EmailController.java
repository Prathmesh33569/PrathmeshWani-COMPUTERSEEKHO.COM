package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.EmailRequestDTO;
import com.computerseekho.api.service.EmailService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/email")
@RequiredArgsConstructor
@CrossOrigin(origins = "*")
public class EmailController {

    private final EmailService emailService;

    @PostMapping
    public ResponseEntity<String> sendEmail(
            @RequestBody EmailRequestDTO dto
    ) {
        emailService.sendEnquiryEmail(dto);
        return ResponseEntity.ok("Email sent successfully");
    }


//    @PostMapping(
//            value = "/receipt/pdf",
//            consumes = "multipart/form-data"
//    )
//    public ResponseEntity<String> sendReceiptPdf(
//            @RequestParam("userEmail") String userEmail,
//            @RequestParam("receiptPdf") MultipartFile receiptPdf
//    ) {
//        emailService.sendReceiptPdfEmail(userEmail, receiptPdf);
//        return ResponseEntity.ok("Receipt PDF sent to user email");
//    }
}
