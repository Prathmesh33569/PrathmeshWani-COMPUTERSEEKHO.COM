package com.computerseekho.api.dto.response;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class StaffJwtResponse {

    private String token;
    private String type = "Bearer";
    private Integer staffId;
    private String staffUsername;
    private String staffEmail;
    private String staffName;
    private String staffRole;

    public StaffJwtResponse(String token, Integer staffId, String staffUsername,
                            String staffEmail, String staffName, String staffRole) {
        this.token = token;
        this.staffId = staffId;
        this.staffUsername = staffUsername;
        this.staffEmail = staffEmail;
        this.staffName = staffName;
        this.staffRole = staffRole;
    }
}