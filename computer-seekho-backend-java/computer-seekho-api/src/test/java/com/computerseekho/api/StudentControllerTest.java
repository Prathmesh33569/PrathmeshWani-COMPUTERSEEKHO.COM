package com.computerseekho.api;

import com.computerseekho.api.controller.StudentController;
import com.computerseekho.api.dto.request.StudentRegistrationDTO;
import com.computerseekho.api.dto.response.StudentResponseDTO;
import com.computerseekho.api.entity.Student;
import com.computerseekho.api.service.StudentService;
import org.junit.jupiter.api.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import static org.mockito.Mockito.when;
import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.extension.ExtendWith;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;

import java.time.LocalDate;

@ExtendWith(MockitoExtension.class)
public class StudentControllerTest {

    @Mock
    private StudentService studentService;

    @InjectMocks
    private StudentController studentController;

    @Test
    public void testRegisterStudent() {


        StudentRegistrationDTO dto = new StudentRegistrationDTO();
        dto.setEnquiryId(1);
        dto.setBatchId(101);

        StudentResponseDTO responseDTO = new StudentResponseDTO();
        responseDTO.setStudentId(1);
        responseDTO.setStudentName("Sachin Avhad");
        responseDTO.setStudentMobile("9876543210");
        responseDTO.setStudentGender("Male");
        responseDTO.setStudentDob(LocalDate.of(2001, 5, 10));
        responseDTO.setStudentQualification("B.Tech");
        responseDTO.setStudentAddress("Pune, Maharashtra");
        responseDTO.setPhotoUrl("http://example.com/photo.jpg");

        responseDTO.setBatchId(101);
        responseDTO.setBatchName("Java Full Stack Batch");

        responseDTO.setCourseId(201);
        responseDTO.setCourseName("Full Stack Java Developer");

        responseDTO.setStudentUsername("sachin123");
        responseDTO.setEnquiryId(1);

        when(studentService.registerStudentFromEnquiry(dto)).thenReturn(responseDTO);


        ResponseEntity<?> response = studentController.registerStudent(dto);


        assertNotNull(response);
        assertEquals(HttpStatus.CREATED, response.getStatusCode());

        StudentResponseDTO result = (StudentResponseDTO) response.getBody();

        assertNotNull(result);
        assertEquals(1, result.getStudentId());
        assertEquals("Sachin Avhad", result.getStudentName());
        assertEquals("9876543210", result.getStudentMobile());
        assertEquals("Male", result.getStudentGender());
        assertEquals(LocalDate.of(2001, 5, 10), result.getStudentDob());
        assertEquals("B.Tech", result.getStudentQualification());
        assertEquals("Pune, Maharashtra", result.getStudentAddress());
        assertEquals("http://example.com/photo.jpg", result.getPhotoUrl());

        assertEquals(101, result.getBatchId());
        assertEquals("Java Full Stack Batch", result.getBatchName());

        assertEquals(201, result.getCourseId());
        assertEquals("Full Stack Java Developer", result.getCourseName());

        assertEquals("sachin123", result.getStudentUsername());
        assertEquals(1, result.getEnquiryId());
    }
}
