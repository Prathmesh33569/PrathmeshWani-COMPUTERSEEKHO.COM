package com.computerseekho.api.controller;

import com.computerseekho.api.dto.response.PaymentPdfDTO;
import com.computerseekho.api.pdf.ReceiptPDFExporter;
import com.computerseekho.api.service.EmailService;
import com.computerseekho.api.service.PaymentService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/receipt")
@RequiredArgsConstructor
@CrossOrigin(origins = "*")
public class ReceiptController {

    private final PaymentService paymentService;

    private final EmailService emailService;


//    @GetMapping("/pdf/{receiptId}")
//    public void generateReceiptPdf(@PathVariable Integer receiptId, HttpServletResponse response) throws Exception {
//
//        response.setContentType("application/pdf");
//        response.setHeader("Content-Disposition", "attachment; filename=receipt.pdf");
//
//        PaymentPdfDTO data = paymentService.getReceiptPdfData(receiptId);
//
//        ReceiptPDFExporter exporter = new ReceiptPDFExporter(data);
//        exporter.export(response);
//    }


    @GetMapping("/pdf/{receiptId}/email")
    public String generatePdfAndSendEmail(@PathVariable Integer receiptId) throws Exception {


        PaymentPdfDTO data = paymentService.getReceiptPdfData(receiptId);
        System.out.println("Student Email: " + data.getStudentEmail());

        ReceiptPDFExporter exporter = new ReceiptPDFExporter(data);

        // Generate PDF in memory
        byte[] pdfBytes = exporter.generatePdf();

        // Get email from DB
        String userEmail = data.getStudentEmail(); // must exist in DTO

        emailService.sendReceiptPdfEmail(userEmail, pdfBytes);

        return "Receipt PDF sent to email: " + userEmail;
    }

}
