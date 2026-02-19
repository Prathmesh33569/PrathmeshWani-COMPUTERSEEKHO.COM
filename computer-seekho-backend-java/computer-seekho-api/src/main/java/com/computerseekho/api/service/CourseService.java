package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.CourseCreateRequestDTO;
import com.computerseekho.api.dto.request.CourseUpdateRequestDTO;
import com.computerseekho.api.dto.response.CourseResponseDTO;
import com.computerseekho.api.entity.Course;

import java.util.List;

public interface CourseService {

    CourseResponseDTO createCourse(CourseCreateRequestDTO dto);

    CourseResponseDTO updateCourse(Integer courseId, CourseUpdateRequestDTO dto);

    List<CourseResponseDTO> getAllCourses();

    List<CourseResponseDTO> getActiveCourses();

    CourseResponseDTO getCourseById(Integer courseId);
    Course getCourseEntityById(Integer id);
    void deleteCourse(Integer courseId);
}
