//package com.computerseekho.api.controller;
//
//import com.computerseekho.api.dto.request.CourseCreateRequestDTO;
//import com.computerseekho.api.dto.request.CourseUpdateRequestDTO;
//import com.computerseekho.api.dto.response.CourseResponseDTO;
//import com.computerseekho.api.service.CourseService;
//import lombok.RequiredArgsConstructor;
//import org.springframework.http.HttpStatus;
//import org.springframework.http.ResponseEntity;
//import org.springframework.web.bind.annotation.*;
//
//import java.util.List;
//
//@RestController
//@RequestMapping("/api/courses")
//@RequiredArgsConstructor
//@CrossOrigin(origins = "*")
//public class CourseController {
//
//    private final CourseService courseService;
//
//    // Add Course
//    @PostMapping
//    public CourseResponseDTO createCourse(@RequestBody CourseCreateRequestDTO dto) {
//        return courseService.createCourse(dto);
//    }
//
//    // Update Course
//    @PutMapping("/{courseId}")
//    public CourseResponseDTO updateCourse(
//            @PathVariable Integer courseId,
//            @RequestBody CourseUpdateRequestDTO dto) {
//        return courseService.updateCourse(courseId, dto);
//    }
//
//    // List all courses
//    @GetMapping
//    public List<CourseResponseDTO> getAllCourses() {
//        return courseService.getAllCourses();
//    }
//
//    // Active courses (used in Enquiry dropdown)
//    @GetMapping("/active")
//    public List<CourseResponseDTO> getActiveCourses() {
//        return courseService.getActiveCourses();
//    }
//
//    // Get course by ID
//    @GetMapping("/{courseId}")
//    public CourseResponseDTO getCourseById(@PathVariable Integer courseId) {
//        return courseService.getCourseById(courseId);
//    }
//
//    @DeleteMapping("/{courseId}")
//    public ResponseEntity<Void> deleteCourse(@PathVariable Integer courseId){
//        courseService.deleteCourse(courseId);
//        return new ResponseEntity<>(HttpStatus.NO_CONTENT);
//    }
//    // delete will be added (task pending for akram assign by sachin)
//}

package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.CourseCreateRequestDTO;
import com.computerseekho.api.dto.request.CourseUpdateRequestDTO;
import com.computerseekho.api.dto.response.CourseResponseDTO;
import com.computerseekho.api.service.CourseService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/courses")
@RequiredArgsConstructor
@CrossOrigin(origins = "*")
public class CourseController {

    private final CourseService courseService;

    // Add Course
    @PostMapping
    public CourseResponseDTO createCourse(@RequestBody CourseCreateRequestDTO dto) {
        return courseService.createCourse(dto);
    }

    // Update Course
    @PutMapping("/{courseId}")
    public CourseResponseDTO updateCourse(
            @PathVariable Integer courseId,
            @RequestBody CourseUpdateRequestDTO dto) {
        return courseService.updateCourse(courseId, dto);
    }

    // List all courses
    @GetMapping
    public List<CourseResponseDTO> getAllCourses() {
        return courseService.getAllCourses();
    }

    // Active courses (used in Enquiry dropdown)
    @GetMapping("/active")
    public List<CourseResponseDTO> getActiveCourses() {
        return courseService.getActiveCourses();
    }

    // Get course by ID
    @GetMapping("/{courseId}")
    public CourseResponseDTO getCourseById(@PathVariable Integer courseId) {
        return courseService.getCourseById(courseId);
    }

    @DeleteMapping("/{courseId}")
    public ResponseEntity<Void> deleteCourse(@PathVariable Integer courseId){
        courseService.deleteCourse(courseId);
        return new ResponseEntity<>(HttpStatus.NO_CONTENT);
    }
    // delete will be added (task pending for akram assign by sachin)
}
