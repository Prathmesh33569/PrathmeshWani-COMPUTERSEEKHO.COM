package com.computerseekho.api.pdf;

import com.computerseekho.api.dto.response.PaymentPdfDTO;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import jakarta.servlet.http.HttpServletResponse;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;

public class ReceiptPDFExporter {

    private PaymentPdfDTO data;

    public ReceiptPDFExporter(PaymentPdfDTO data) {
        this.data = data;
    }

    // ✅ 1) Existing Method - Download PDF
    public void export(HttpServletResponse response) throws DocumentException, IOException {

        Document document = new Document(PageSize.A4);
        PdfWriter.getInstance(document, response.getOutputStream());
        document.open();

        buildPdf(document);

        document.close();
    }

    // ✅ 2) New Method - Generate PDF as byte[] (for Email)
    public byte[] generatePdf() throws Exception {

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Document document = new Document(PageSize.A4);
        PdfWriter.getInstance(document, outputStream);
        document.open();

        buildPdf(document);

        document.close();

        return outputStream.toByteArray();
    }

    // ✅ Common PDF Content (Reusable)
    private void buildPdf(Document document) throws DocumentException {

        Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16);
        Paragraph title = new Paragraph("PAYMENT RECEIPT\n\n", titleFont);
        title.setAlignment(Paragraph.ALIGN_CENTER);
        document.add(title);

        PdfPTable table = new PdfPTable(2);
        table.setWidthPercentage(100);
        table.setSpacingBefore(10);

        addRow(table, "Student Name", data.getStudentName());
        addRow(table, "Mobile", data.getStudentMobile());
        addRow(table, "Address", data.getStudentAddress());
        addRow(table, "Course", data.getCourseName());
        addRow(table, "Payment Type", data.getPaymentType());
        addRow(table, "Amount", String.valueOf(data.getAmount()));

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");

        addRow(table, "Payment Date",
                data.getPaymentDate() != null ? data.getPaymentDate().format(formatter) : "");

        addRow(table, "Receipt Amount", String.valueOf(data.getReceiptAmount()));

        addRow(table, "Receipt Date",
                data.getReceiptDate() != null ? data.getReceiptDate().format(formatter) : "");

        document.add(table);
    }

    private void addRow(PdfPTable table, String label, String value) {
        PdfPCell cell1 = new PdfPCell(new Phrase(label));
        cell1.setPadding(5);

        PdfPCell cell2 = new PdfPCell(new Phrase(value != null ? value : ""));
        cell2.setPadding(5);

        table.addCell(cell1);
        table.addCell(cell2);
    }
}
