//package com.computerseekho.api.dto.response;
//
//import lombok.Data;
//
//import java.math.BigDecimal;
//import java.time.LocalDate;
//
//@Data
//public class CourseResponseDTO {
//
//    private Integer courseId;
//    private String courseName;
//    private String courseDescription;
//    private Integer courseDuration;
//    private BigDecimal courseFees;
//    private LocalDate courseFeesFrom;      // Add this
//    private LocalDate courseFeesTo;        // Add this
//    private String courseSyllabus;         // Add this
//    private Boolean courseIsActive;
//    private String ageGrpType;
//    private String coverPhoto;             // ← ADD THIS!
//    private Integer videoId;               // Add this if needed
//}

package com.computerseekho.api.dto.response;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class CourseResponseDTO {

    private Integer courseId;
    private String courseName;
    private String courseDescription;
    private Integer courseDuration;
    private BigDecimal courseFees;
    private LocalDate courseFeesFrom;      // Add this
    private LocalDate courseFeesTo;        // Add this
    private String courseSyllabus;         // Add this
    private Boolean courseIsActive;
    private String ageGrpType;
    private String coverPhoto;             // ← ADD THIS!
    private Integer videoId;               // Add this if needed
}