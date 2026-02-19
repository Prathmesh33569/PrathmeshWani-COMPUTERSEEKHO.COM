//package com.computerseekho.api.dto.response;
//
//import lombok.AllArgsConstructor;
//import lombok.Data;
//import lombok.NoArgsConstructor;
//
//import java.math.BigDecimal;
//import java.time.LocalDate;
//import java.time.LocalDateTime;
//
//@Data
//@AllArgsConstructor
//@NoArgsConstructor
//// REMOVED: @JsonNaming(PropertyNamingStrategies.SnakeCaseStrategy.class)
//public class BatchResponseDTO {
//    private Integer batchId;
//    private String batchName;
//    private LocalDate batchStartDate;
//    private LocalDate batchEndDate;
//    private Integer courseId;
//    private String courseName;
//    private LocalDateTime presentationDate;
//    private BigDecimal courseFees;
//    private LocalDate courseFeesFrom;
//    private LocalDate courseFeesTo;
//    private Boolean batchIsActive;
//    private LocalDateTime finalPresentationDate;
//    private String batchLogoUrl;
//    private LocalDateTime createdAt;
//    private LocalDateTime updatedAt;
//}

package com.computerseekho.api.dto.response;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

@Data
@AllArgsConstructor
@NoArgsConstructor
// REMOVED: @JsonNaming(PropertyNamingStrategies.SnakeCaseStrategy.class)
public class BatchResponseDTO {
    private Integer batchId;
    private String batchName;
    private LocalDate batchStartDate;
    private LocalDate batchEndDate;
    private Integer courseId;
    private String courseName;
    private LocalDateTime presentationDate;
    private BigDecimal courseFees;
    private LocalDate courseFeesFrom;
    private LocalDate courseFeesTo;
    private Boolean batchIsActive;
    private LocalDateTime finalPresentationDate;
    private String batchLogoUrl;
    private LocalDateTime createdAt;
    private LocalDateTime updatedAt;
}