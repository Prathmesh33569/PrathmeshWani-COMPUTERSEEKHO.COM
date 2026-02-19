package com.computerseekho.api.dto.request;

import jakarta.validation.constraints.Email;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Size;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class StaffSignupRequest {

    @NotBlank(message = "Staff name is required")
    @Size(max = 100, message = "Staff name must be less than 100 characters")
    private String staffName;

    private String photoUrl;

    @NotBlank(message = "Mobile number is required")
    private String staffMobile;

    @NotBlank(message = "Email is required")
    @Email(message = "Email should be valid")
    @Size(max = 100)
    private String staffEmail;

    @NotBlank(message = "Username is required")
    @Size(min = 3, max = 100, message = "Username must be between 3 and 100 characters")
    private String staffUsername;

    @NotBlank(message = "Password is required")
    @Size(min = 6, max = 40, message = "Password must be between 6 and 40 characters")
    private String staffPassword;

    @NotBlank(message = "Staff role is required")
    private String staffRole;  // "teaching" or "non-teaching"

    private String staffBio;

    private String staffDesignation;
}