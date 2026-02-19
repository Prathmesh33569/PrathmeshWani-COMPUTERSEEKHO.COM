package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.EmailRequestDTO;
import jakarta.mail.internet.MimeMessage;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
@RequiredArgsConstructor
public class EmailService {

    private final JavaMailSender mailSender;

    public void sendEnquiryEmail(EmailRequestDTO dto) {

        SimpleMailMessage message = new SimpleMailMessage();

        message.setTo("sheekhocomputer@gmail.com");   // FIXED RECEIVER
        message.setSubject("New Enquiry Received");

        message.setText(
                "Name : " + dto.getName() + "\n" +
                        "Email: " + dto.getEmail() + "\n\n" +
                        "Message:\n" + dto.getMessage()
        );

        mailSender.send(message);
    }


//    public void sendReceiptPdfEmail(
//            String userEmail,
//            MultipartFile receiptPdf
//    ) {
//        try {
//            MimeMessage message = mailSender.createMimeMessage();
//
//            MimeMessageHelper helper =
//                    new MimeMessageHelper(message, true);
//
//            helper.setTo(userEmail);
//            helper.setSubject("Payment Receipt - Computer Seekho");
//
//            helper.setText(
//                    "Hello,\n\n" +
//                            "Please find your payment receipt attached as a PDF.\n\n" +
//                            "Thank you for choosing Computer Seekho.\n\n" +
//                            "Regards,\n" +
//                            "Computer Seekho Team"
//            );
//
//            helper.addAttachment(
//                    receiptPdf.getOriginalFilename(),
//                    receiptPdf
//            );
//
//            mailSender.send(message);
//
//        } catch (Exception e) {
//            throw new RuntimeException(
//                    "Failed to send receipt PDF email",
//                    e
//            );
//        }


    public void sendReceiptPdfEmail(String userEmail, byte[] pdfBytes) {

        try {
            MimeMessage message = mailSender.createMimeMessage();
            MimeMessageHelper helper = new MimeMessageHelper(message, true);

            helper.setTo(userEmail);
            helper.setSubject("Payment Receipt - Computer Seekho");

            helper.setText(
                    "Hello,\n\n" +
                            "Please find your payment receipt attached as a PDF.\n\n" +
                            "Regards,\nComputer Seekho Team"
            );

            helper.addAttachment("receipt.pdf", new ByteArrayResource(pdfBytes));

            mailSender.send(message);

        } catch (Exception e) {
            e.printStackTrace();  // 🔥 Add this line
            throw new RuntimeException("Failed to send receipt PDF email", e);
        }
    }

}

