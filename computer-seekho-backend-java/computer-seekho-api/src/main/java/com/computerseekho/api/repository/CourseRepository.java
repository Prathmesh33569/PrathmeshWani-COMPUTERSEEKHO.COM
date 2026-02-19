package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Course;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface CourseRepository extends JpaRepository<Course, Integer> {

    // Active courses only (used in enquiry dropdown)
    List<Course> findByCourseIsActiveTrue();
}

