package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.StaffLoginRequest;
import com.computerseekho.api.dto.response.MessageResponse;
import com.computerseekho.api.dto.response.StaffJwtResponse;
import com.computerseekho.api.service.StaffAuthService;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.validation.Valid;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import org.springframework.security.core.annotation.AuthenticationPrincipal;
import org.springframework.security.oauth2.core.user.OAuth2User;

import java.io.IOException;

@CrossOrigin(origins = "*", maxAge = 3600)
@RestController
@RequestMapping("/api/auth")
public class StaffAuthController {

    @Autowired
    private StaffAuthService staffAuthService;

    /**
     * Staff Login endpoint
     * POST /api/auth/login
     */
    @PostMapping("/login")
    public ResponseEntity<?> authenticateStaff(@Valid @RequestBody StaffLoginRequest loginRequest) {
        StaffJwtResponse jwtResponse = staffAuthService.authenticateStaff(loginRequest);
        return ResponseEntity.ok(jwtResponse);
    }

    /**
     * Staff Signup/Register endpoint with IMAGE UPLOAD
     * POST /api/auth/signup
     *
     * Accepts multipart/form-data with:
     * - staffImage (file, optional)
     * - All other staff fields as form data
     */
    @PostMapping("/signup")
    public ResponseEntity<?> registerStaff(
            @RequestParam(value = "staffImage", required = false) MultipartFile staffImage,
            @RequestParam("staffName") String staffName,
            @RequestParam("staffMobile") String staffMobile,
            @RequestParam("staffEmail") String staffEmail,
            @RequestParam("staffUsername") String staffUsername,
            @RequestParam("staffPassword") String staffPassword,
            @RequestParam("staffRole") String staffRole,
            @RequestParam(value = "staffDesignation", required = false) String staffDesignation,
            @RequestParam(value = "staffBio", required = false) String staffBio
    ) {
        try {
            MessageResponse messageResponse = staffAuthService.registerStaffWithImage(
                    staffImage, staffName, staffMobile, staffEmail,
                    staffUsername, staffPassword, staffRole,
                    staffDesignation, staffBio
            );

            // Check if registration was successful
            if (messageResponse.getMessage().startsWith("Error")) {
                return ResponseEntity.badRequest().body(messageResponse);
            }

            return ResponseEntity.ok(messageResponse);

        } catch (Exception e) {
            return ResponseEntity.badRequest()
                    .body(new MessageResponse("Error: " + e.getMessage()));
        }
    }

    @GetMapping("/oauth2/user")
    public void getOAuth2User(@AuthenticationPrincipal OAuth2User oAuth2User,
                              HttpServletResponse response) throws IOException {

        if (oAuth2User == null) {
            response.sendRedirect("http://localhost:5173/admin/login?error=oauth");
            return;
        }

        String email = oAuth2User.getAttribute("email");

        // Generate JWT using Google email
        StaffJwtResponse jwtResponse = staffAuthService.authenticateStaffWithGoogle(email);

        String token = jwtResponse.getToken();

        // Redirect to React with token
        response.sendRedirect("http://localhost:5173/oauth-success?token=" + token);
    }

    /**
     * Test endpoint to check if API is working
     * GET /api/auth/test
     */
    @GetMapping("/test")
    public ResponseEntity<?> test() {
        return ResponseEntity.ok(new MessageResponse("Staff Auth API is working!"));
    }
}