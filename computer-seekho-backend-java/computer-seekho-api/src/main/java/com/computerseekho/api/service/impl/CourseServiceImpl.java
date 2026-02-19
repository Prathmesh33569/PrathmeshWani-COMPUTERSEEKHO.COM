package com.computerseekho.api.service.impl;

import com.computerseekho.api.dto.request.CourseCreateRequestDTO;
import com.computerseekho.api.dto.request.CourseUpdateRequestDTO;
import com.computerseekho.api.dto.response.CourseResponseDTO;
import com.computerseekho.api.entity.Course;
import com.computerseekho.api.repository.CourseRepository;
import com.computerseekho.api.service.CourseService;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class CourseServiceImpl implements CourseService {

    private final CourseRepository courseRepository;

    @Override
    public CourseResponseDTO createCourse(CourseCreateRequestDTO dto) {

        Course course = new Course();
        course.setCourseName(dto.getCourseName());
        course.setCourseDescription(dto.getCourseDescription());
        course.setCourseDuration(dto.getCourseDuration());
        course.setCourseFees(dto.getCourseFees());
        course.setCourseFeesFrom(dto.getCourseFeesFrom());
        course.setCourseFeesTo(dto.getCourseFeesTo());
        course.setCourseSyllabus(dto.getCourseSyllabus());
        course.setAgeGrpType(dto.getAgeGrpType());
        course.setCoverPhoto(dto.getCoverPhoto());
        course.setVideoId(dto.getVideoId());

        return mapToDTO(courseRepository.save(course));
    }

    @Override
    public CourseResponseDTO updateCourse(Integer courseId, CourseUpdateRequestDTO dto) {

        Course course = courseRepository.findById(courseId).orElseThrow();

        course.setCourseName(dto.getCourseName());
        course.setCourseDescription(dto.getCourseDescription());
        course.setCourseDuration(dto.getCourseDuration());
        course.setCourseFees(dto.getCourseFees());
        course.setCourseFeesFrom(dto.getCourseFeesFrom());
        course.setCourseFeesTo(dto.getCourseFeesTo());
        course.setCourseSyllabus(dto.getCourseSyllabus());
        course.setAgeGrpType(dto.getAgeGrpType());
        course.setCourseIsActive(dto.getCourseIsActive());
        course.setCoverPhoto(dto.getCoverPhoto());
        course.setVideoId(dto.getVideoId());

        return mapToDTO(courseRepository.save(course));
    }

    @Override
    public List<CourseResponseDTO> getAllCourses() {
        return courseRepository.findAll()
                .stream()
                .map(this::mapToDTO)
                .collect(Collectors.toList());
    }

    @Override
    public List<CourseResponseDTO> getActiveCourses() {
        return courseRepository.findByCourseIsActiveTrue()
                .stream()
                .map(this::mapToDTO)
                .collect(Collectors.toList());
    }

    @Override
    public CourseResponseDTO getCourseById(Integer courseId) {
        return mapToDTO(courseRepository.findById(courseId).orElseThrow());
    }

    private CourseResponseDTO mapToDTO(Course course) {
        CourseResponseDTO dto = new CourseResponseDTO();
        dto.setCourseId(course.getCourseId());
        dto.setCourseDescription((course.getCourseDescription()));
        dto.setCourseName(course.getCourseName());
        dto.setCourseDuration(course.getCourseDuration());
        dto.setCourseFees(course.getCourseFees());
        dto.setCourseIsActive(course.getCourseIsActive());
        dto.setAgeGrpType(course.getAgeGrpType());
        dto.setCoverPhoto(course.getCoverPhoto());
        dto.setCourseSyllabus(course.getCourseSyllabus());
        return dto;
    }
    @Override
    public Course getCourseEntityById(Integer id) {
        return courseRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Course not found with id: " + id));
    }

    @Override
    public void deleteCourse(Integer courseId){
        courseRepository.deleteById(courseId);
    }
}
