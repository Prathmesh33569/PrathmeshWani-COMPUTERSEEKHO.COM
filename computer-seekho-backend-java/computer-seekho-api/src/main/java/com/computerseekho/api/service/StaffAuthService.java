package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.StaffLoginRequest;
import com.computerseekho.api.dto.response.MessageResponse;
import com.computerseekho.api.dto.response.StaffJwtResponse;
import com.computerseekho.api.entity.Staff;
import com.computerseekho.api.repository.StaffRepository;
import com.computerseekho.api.security.jwt.JwtUtils;
import com.computerseekho.api.security.services.StaffDetailsImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class StaffAuthService {

    @Autowired
    private AuthenticationManager authenticationManager;

    @Autowired
    private StaffRepository staffRepository;

    @Autowired
    private PasswordEncoder passwordEncoder;

    @Autowired
    private JwtUtils jwtUtils;

    @Autowired
    private FileStorageServices fileStorageServices;

    // Login staff
    public StaffJwtResponse authenticateStaff(StaffLoginRequest loginRequest) {
        // Authenticate staff
        Authentication authentication = authenticationManager.authenticate(
                new UsernamePasswordAuthenticationToken(
                        loginRequest.getStaffUsername(),
                        loginRequest.getStaffPassword()
                )
        );

        // Set authentication in security context
        SecurityContextHolder.getContext().setAuthentication(authentication);

        // Generate JWT token
        String jwt = jwtUtils.generateJwtToken(authentication);

        // Get staff details
        StaffDetailsImpl staffDetails = (StaffDetailsImpl) authentication.getPrincipal();

        // Return JWT response
        return new StaffJwtResponse(
                jwt,
                staffDetails.getStaffId(),
                staffDetails.getStaffUsername(),
                staffDetails.getStaffEmail(),
                staffDetails.getStaffName(),
                staffDetails.getStaffRole()
        );
    }

//    // Register new staff
//    public MessageResponse registerStaff(StaffSignupRequest signupRequest) {
//        // Check if username already exists
//        if (staffRepository.existsByStaffUsername(signupRequest.getStaffUsername())) {
//            return new MessageResponse("Error: Username is already taken!");
//        }
//
//        // Check if email already exists
//        if (staffRepository.existsByStaffEmail(signupRequest.getStaffEmail())) {
//            return new MessageResponse("Error: Email is already in use!");
//        }
//
//        // Validate staff role
//        String role = signupRequest.getStaffRole().toLowerCase();
//        if (!role.equals("teaching") && !role.equals("non-teaching")) {
//            return new MessageResponse("Error: Invalid staff role! Use 'teaching' or 'non-teaching'");
//        }
//
//        // Create new staff account
//        Staff staff = new Staff(
//                signupRequest.getStaffUsername(),
//                signupRequest.getStaffEmail(),
//                passwordEncoder.encode(signupRequest.getStaffPassword())  // Encrypt password
//        );
//
//        // Set additional fields
//        staff.setStaffName(signupRequest.getStaffName());
//        staff.setPhotoUrl(signupRequest.getPhotoUrl());
//        staff.setStaffMobile(signupRequest.getStaffMobile());
//        staff.setStaffRole(role);
//        staff.setStaffBio(signupRequest.getStaffBio());
//        staff.setStaffDesignation(signupRequest.getStaffDesignation());
//
//        staffRepository.save(staff);
//
//        return new MessageResponse("Staff registered successfully!");
//    }

    /**
     * Register new staff WITH IMAGE UPLOAD
     */
    public MessageResponse registerStaffWithImage(
            MultipartFile staffImage,
            String staffName,
            String staffMobile,
            String staffEmail,
            String staffUsername,
            String staffPassword,
            String staffRole,
            String staffDesignation,
            String staffBio
    ) {
        // Check if username already exists
        if (staffRepository.existsByStaffUsername(staffUsername)) {
            return new MessageResponse("Error: Username is already taken!");
        }

        // Check if email already exists
        if (staffRepository.existsByStaffEmail(staffEmail)) {
            return new MessageResponse("Error: Email is already in use!");
        }

        // Validate staff role
        String role = staffRole.toLowerCase();
        if (!role.equals("teaching") && !role.equals("non-teaching")) {
            return new MessageResponse("Error: Invalid staff role! Use 'teaching' or 'non-teaching'");
        }

        // Handle image upload
        String photoUrl = null;
        if (staffImage != null && !staffImage.isEmpty()) {
            // Validate file size (max 5MB)
            if (!fileStorageServices.isValidFileSize(staffImage)) {
                return new MessageResponse("Error: File size exceeds 5MB limit!");
            }

            try {
                // Store image and get URL
                photoUrl = fileStorageServices.storeStaffImage(staffImage, staffUsername);
            } catch (Exception e) {
                return new MessageResponse("Error: Failed to upload image - " + e.getMessage());
            }
        }

        // Create new staff account
        Staff staff = new Staff(
                staffUsername,
                staffEmail,
                passwordEncoder.encode(staffPassword)  // Encrypt password
        );

        // Set additional fields
        staff.setStaffName(staffName);
        staff.setPhotoUrl(photoUrl);  // Store image path
        staff.setStaffMobile(staffMobile);
        staff.setStaffRole(role);
        staff.setStaffBio(staffBio);
        staff.setStaffDesignation(staffDesignation);

        staffRepository.save(staff);

        return new MessageResponse("Staff registered successfully!");
    }

    public StaffJwtResponse authenticateStaffWithGoogle(String email) {

        Staff staff = staffRepository.findByStaffEmail(email)
                .orElseThrow(() -> new RuntimeException("Staff not found with email: " + email));


        String jwt = jwtUtils.generateTokenFromUsername(staff.getStaffUsername());

        return new StaffJwtResponse(
                jwt,
                staff.getStaffId(),
                staff.getStaffUsername(),
                staff.getStaffEmail(),
                staff.getStaffName(),
                staff.getStaffRole()
        );
    }

}