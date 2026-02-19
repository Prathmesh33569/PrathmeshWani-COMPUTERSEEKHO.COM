//package com.computerseekho.api.dto.request;
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
//public class BatchCreateRequestDTO {
//    private String batchName;
//    private LocalDate batchStartDate;
//    private LocalDate batchEndDate;
//    private Integer courseId;
//    private LocalDateTime presentationDate;
//    private BigDecimal courseFees;
//    private LocalDate courseFeesFrom;
//    private LocalDate courseFeesTo;
//    private Boolean batchIsActive;
//    private LocalDateTime finalPresentationDate;
//    private String batchLogoUrl;
//}

package com.computerseekho.api.dto.request;

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
public class BatchCreateRequestDTO {
    private String batchName;
    private LocalDate batchStartDate;
    private LocalDate batchEndDate;
    private Integer courseId;
    private LocalDateTime presentationDate;
    private BigDecimal courseFees;
    private LocalDate courseFeesFrom;
    private LocalDate courseFeesTo;
    private Boolean batchIsActive;
    private LocalDateTime finalPresentationDate;
    private String batchLogoUrl;
}