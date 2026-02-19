package com.computerseekho.api.dto.response;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class StaffDTO {
    private Integer staffId;
    private String staffName;
    private String photoUrl;
    private String staffMobile;
    private String staffEmail;
    private String staffRole;
    private String staffBio;
    private String staffDesignation;
    private Boolean isActive;
    private LocalDateTime createdAt;
    private LocalDateTime updatedAt;
}
