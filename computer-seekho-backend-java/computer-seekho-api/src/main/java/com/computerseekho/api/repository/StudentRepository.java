package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Student;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface StudentRepository extends JpaRepository<Student,Integer> {
    @Query("SELECT s FROM Student s WHERE s.studentId IN " +
            "(SELECT p.studentId FROM Payment p WHERE p.studentId IS NOT NULL " +
            "AND p.paymentId IN (SELECT r.paymentId FROM Receipt r))")
    List<Student> findAllRegisteredStudentsWithPaymentAndReceipt();
}
