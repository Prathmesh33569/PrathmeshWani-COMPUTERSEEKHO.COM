package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.StudentRegistrationDTO;
import com.computerseekho.api.dto.response.RegisteredStudentResponseDTO;
import com.computerseekho.api.dto.response.StudentResponseDTO;
import com.computerseekho.api.entity.Student;
import com.computerseekho.api.service.StudentService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/students")
@RequiredArgsConstructor
@CrossOrigin(origins = "*")
public class StudentController {

    private final StudentService studentService;

    /**
     * Register student from enquiry (AFTER payment)
     * POST /api/students/register
     */
    @PostMapping("/register")
    public ResponseEntity<?> registerStudent(@RequestBody StudentRegistrationDTO dto) {
        try {
            StudentResponseDTO student = studentService.registerStudentFromEnquiry(dto);
            return ResponseEntity.status(HttpStatus.CREATED).body(student);
        } catch (RuntimeException e) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(e.getMessage());
        }
    }

    /**
     * Get all registered students with complete details (payment & receipt)
     * GET /api/students/registered
     */
    @GetMapping("/registered")
    public ResponseEntity<List<RegisteredStudentResponseDTO>> getAllRegisteredStudents() {
        List<RegisteredStudentResponseDTO> students = studentService.getAllRegisteredStudentsWithDetails();
        return ResponseEntity.ok(students);
    }

    /**
     * Get student by ID
     * GET /api/students/{id}
     */

    @GetMapping("/id/{id}")
    public ResponseEntity<?> getStudentById(@PathVariable Integer id) {
        try {
            Student student = studentService.getStudentById(id);
            return ResponseEntity.ok(student);
        } catch (RuntimeException e) {
            return ResponseEntity.status(HttpStatus.NOT_FOUND).body(e.getMessage());
        }
    }


    /**
     * Get all students
     * GET /api/students
     */
    @GetMapping
    public ResponseEntity<List<StudentResponseDTO>> getAllStudents() {
        List<StudentResponseDTO> students = studentService.getAllStudents();
        return ResponseEntity.ok(students);
    }

    /**
     * Get students by course
     * GET /api/students/course/{courseId}
     */
    @GetMapping("/course/{courseId}")
    public ResponseEntity<List<StudentResponseDTO>> getStudentsByCourse(@PathVariable Integer courseId) {
        List<StudentResponseDTO> students = studentService.getStudentsByCourse(courseId);
        return ResponseEntity.ok(students);
    }
}