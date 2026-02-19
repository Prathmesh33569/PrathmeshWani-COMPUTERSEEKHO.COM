package com.computerseekho.api.repository;

import com.computerseekho.api.dto.response.PaymentPdfDTO;
import com.computerseekho.api.entity.Payment;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface PaymentRepository extends JpaRepository<Payment, Integer> {

    List<Payment> findByStudentId(Integer studentId);

    List<Payment> findByCourseId(Integer courseId);

    List<Payment> findByBatchId(Integer batchId);

    Optional<Payment> findByEnquiryId(Integer enquiryId);

    @Query(value = """
    SELECT 
        s.student_name,
        s.student_mobile,
        s.student_address,
        s.student_username,
        c.course_name,
        p.amount,
        p.payment_date,
        pt.payment_type_desc,
        r.receipt_amount,
        r.receipt_date
    FROM receipt r
    JOIN payment p ON r.payment_id = p.payment_id
    JOIN student_master s ON p.student_id = s.student_id
    JOIN course_master c ON p.course_id = c.course_id
    JOIN payment_type_master pt ON p.payment_type_id = pt.payment_type_id
    WHERE r.receipt_id = :receiptId
    """, nativeQuery = true)
    List<Object[]> getReceiptPdfData(@Param("receiptId") Integer receiptId);



}