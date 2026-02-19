//package com.computerseekho.api.dto.request;
//
//import lombok.Data;
//
//import java.math.BigDecimal;
//import java.time.LocalDate;
//
//@Data
//public class CourseCreateRequestDTO {
//
//    private String courseName;
//    private String courseDescription;
//    private Integer courseDuration;
//    private BigDecimal courseFees;
//    private LocalDate courseFeesFrom;
//    private LocalDate courseFeesTo;
//    private String courseSyllabus;
//    private String ageGrpType;        // Adult / Kids / All Ages
//    private String coverPhoto;
//    private Integer videoId;
//}
//

package com.computerseekho.api.dto.request;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class CourseCreateRequestDTO {

    private String courseName;
    private String courseDescription;
    private Integer courseDuration;
    private BigDecimal courseFees;
    private LocalDate courseFeesFrom;
    private LocalDate courseFeesTo;
    private String courseSyllabus;
    private String ageGrpType;        // Adult / Kids / All Ages
    private String coverPhoto;
    private Integer videoId;
}


