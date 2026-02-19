package com.computerseekho.api.dto.request;

import jakarta.validation.constraints.NotBlank;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class StaffLoginRequest {

    @NotBlank(message = "Username is required")
    private String staffUsername;

    @NotBlank(message = "Password is required")
    private String staffPassword;
}