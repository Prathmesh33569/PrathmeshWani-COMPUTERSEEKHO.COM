package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.StudentRegistrationDTO;
import com.computerseekho.api.dto.response.RegisteredStudentResponseDTO;
import com.computerseekho.api.dto.response.StudentResponseDTO;
import com.computerseekho.api.entity.Student;

import java.util.List;

public interface StudentService {

    Student saveStudent(Student student);

    // Register student from enquiry
    StudentResponseDTO registerStudentFromEnquiry(StudentRegistrationDTO dto);

    // Get student by ID
    Student getStudentById(Integer studentId);

    // Get all students
    List<StudentResponseDTO> getAllStudents();

    // Get students by course
    List<StudentResponseDTO> getStudentsByCourse(Integer courseId);

    // NEW METHOD
    List<RegisteredStudentResponseDTO> getAllRegisteredStudentsWithDetails();
}