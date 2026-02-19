package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.StudentRegistrationDTO;
import com.computerseekho.api.dto.response.RegisteredStudentResponseDTO;
import com.computerseekho.api.dto.response.StudentResponseDTO;
import com.computerseekho.api.entity.*;
import com.computerseekho.api.repository.*;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class StudentServiceImpl implements StudentService {

    private final StudentRepository studentRepository;
    private final EnquiryRepository enquiryRepository;
    private final PaymentRepository paymentRepository;
    private final ReceiptRepository receiptRepository;
    private final PaymentTypeRepository paymentTypeRepository;

    @Override
    public Student saveStudent(Student student) {
        return studentRepository.save(student);
    }

    @Override
    @Transactional
    public StudentResponseDTO registerStudentFromEnquiry(StudentRegistrationDTO dto) {

        // 1. Validate payment exists
        if (dto.getPaymentId() == null) {
            throw new RuntimeException("Payment ID is required for student registration");
        }

        Payment payment = paymentRepository.findById(dto.getPaymentId())
                .orElseThrow(() -> new RuntimeException("Payment not found with id: " + dto.getPaymentId()));

        // 2. Get the enquiry
        Enquiry enquiry = enquiryRepository.findById(dto.getEnquiryId())
                .orElseThrow(() -> new RuntimeException("Enquiry not found with id: " + dto.getEnquiryId()));

        // 3. Check if enquiry is already processed
        if (Boolean.TRUE.equals(enquiry.getEnquiryProcessedFlag())) {
            throw new RuntimeException("This enquiry has already been converted to student registration");
        }

        // 4. Check if enquiry is closed
        if (Boolean.TRUE.equals(enquiry.getIsClosed())) {
            throw new RuntimeException("Cannot register student from a closed enquiry");
        }

        // 5. Create Student entity
        Student student = new Student();
        student.setStudentName(dto.getStudentName());
        student.setStudentMobile(dto.getStudentMobile());
        student.setStudentAddress(dto.getStudentAddress());
        student.setStudentGender(dto.getStudentGender());
        student.setStudentDob(dto.getStudentDob());
        student.setStudentQualification(dto.getStudentQualification());
        student.setPhotoUrl(dto.getPhotoUrl());
        student.setCourseId(dto.getCourseId());
        student.setBatchId(dto.getBatchId());
        student.setStudentUsername(dto.getStudentUsername());
        student.setStudentPassword(dto.getStudentPassword()); // TODO: Hash password

        // 6. Save student
        Student savedStudent = studentRepository.save(student);

        // 7. Update payment with student ID
        payment.setStudentId(savedStudent.getStudentId());
        paymentRepository.save(payment);

        // 8. Mark enquiry as processed and closed
        enquiry.setEnquiryProcessedFlag(true);
        enquiry.setIsClosed(true);
        enquiryRepository.save(enquiry);

        // 9. Return response DTO
        return mapToResponseDTO(savedStudent, enquiry);
    }

    @Override
    public Student getStudentById(Integer studentId) {
        return studentRepository.findById(studentId)
                .orElseThrow(() -> new RuntimeException("Student not found with id: " + studentId));
    }

    @Override
    public List<StudentResponseDTO> getAllStudents() {
        return studentRepository.findAll().stream()
                .map(student -> mapToResponseDTO(student, null))
                .collect(Collectors.toList());
    }

    @Override
    public List<StudentResponseDTO> getStudentsByCourse(Integer courseId) {
        // You'll need to add this method to StudentRepository
        // return studentRepository.findByCourseId(courseId).stream()
        //         .map(student -> mapToResponseDTO(student, null))
        //         .collect(Collectors.toList());
        return List.of(); // Placeholder
    }

    // ✅ UPDATED: Only fetch students with payment AND receipt
    @Override
    public List<RegisteredStudentResponseDTO> getAllRegisteredStudentsWithDetails() {

        List<Student> students = studentRepository.findAll();

        return students.stream()
                .map(this::mapToRegisteredStudentDTO)
                .filter(dto -> dto != null) // Filter out students without payment/receipt
                .collect(Collectors.toList());
    }

    private StudentResponseDTO mapToResponseDTO(Student student, Enquiry enquiry) {
        StudentResponseDTO dto = new StudentResponseDTO();
        dto.setStudentId(student.getStudentId());
        dto.setStudentName(student.getStudentName());
        dto.setStudentMobile(student.getStudentMobile() != null ? student.getStudentMobile().toString() : null);
        dto.setStudentGender(student.getStudentGender());
        dto.setStudentDob(student.getStudentDob());
        dto.setStudentQualification(student.getStudentQualification());
        dto.setStudentAddress(student.getStudentAddress());
        dto.setPhotoUrl(student.getPhotoUrl());
        dto.setBatchId(student.getBatchId());
        dto.setCourseId(student.getCourseId());
        dto.setStudentUsername(student.getStudentUsername());

        if (enquiry != null) {
            dto.setEnquiryId(enquiry.getEnquiryId());
        }

        return dto;
    }

    // ✅ UPDATED: Returns null if student doesn't have payment or receipt
    private RegisteredStudentResponseDTO mapToRegisteredStudentDTO(Student student) {

        // ✅ CHECK 1: Student must have payment
        List<Payment> payments = paymentRepository.findByStudentId(student.getStudentId());

        if (payments.isEmpty()) {
            return null; // Skip this student - no payment
        }

        Payment payment = payments.get(0);

        // ✅ CHECK 2: Payment must have receipt
        Optional<Receipt> receiptOpt = receiptRepository.findByPaymentId(payment.getPaymentId());

        if (receiptOpt.isEmpty()) {
            return null; // Skip this student - no receipt
        }

        Receipt receipt = receiptOpt.get();

        // ✅ Student is fully registered - create DTO
        RegisteredStudentResponseDTO dto = new RegisteredStudentResponseDTO();

        // Student Info
        dto.setStudentId(student.getStudentId());
        dto.setStudentName(student.getStudentName());
        dto.setStudentMobile(
                student.getStudentMobile() != null ? student.getStudentMobile().toString() : null
        );
        dto.setStudentGender(student.getStudentGender());
        dto.setStudentDob(student.getStudentDob());
        dto.setStudentQualification(student.getStudentQualification());
        dto.setStudentAddress(student.getStudentAddress());
        dto.setPhotoUrl(student.getPhotoUrl());
        dto.setRegisteredAt(student.getCreatedAt());

        // Course & Batch Info
        dto.setCourseId(student.getCourseId());
        dto.setBatchId(student.getBatchId());

        // Payment Info
        dto.setPaymentId(payment.getPaymentId());
        dto.setAmount(payment.getAmount());
        dto.setPaymentDate(payment.getPaymentDate());
        dto.setPaymentTypeId(payment.getPaymentTypeId());
        dto.setEnquiryId(payment.getEnquiryId());

        // Payment Type Description
        paymentTypeRepository.findById(payment.getPaymentTypeId())
                .ifPresent(pt -> dto.setPaymentTypeDesc(pt.getPaymentTypeDesc()));

        // Receipt Info
        dto.setReceiptId(receipt.getReceiptId());
        dto.setReceiptDate(receipt.getReceiptDate());
        dto.setReceiptAmount(receipt.getReceiptAmount());

        return dto;
    }
}