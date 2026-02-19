package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.EnquiryCreateRequestDTO;
import com.computerseekho.api.dto.request.EnquiryFollowUpRequestDTO;
import com.computerseekho.api.dto.request.EnquiryUpdateRequestDTO;
import com.computerseekho.api.dto.response.EnquiryResponseDTO;
import com.computerseekho.api.entity.Enquiry;
import com.computerseekho.api.service.EnquiryService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/enquiries")
@RequiredArgsConstructor
@CrossOrigin(origins = "*")
public class EnquiryController {

    private final EnquiryService enquiryService;

    // Add Enquiry Page
    @PostMapping
    public Enquiry createEnquiry(@RequestBody EnquiryCreateRequestDTO dto) {
        return enquiryService.createEnquiry(dto);
    }

    // Get Single Enquiry by ID (for Edit functionality)
    @GetMapping("/{id}")
    public ResponseEntity<Enquiry> getEnquiryById(@PathVariable Integer id) {
        Enquiry enquiry = enquiryService.getEnquiryById(id);
        return ResponseEntity.ok(enquiry);
    }

    // Update Enquiry (Edit functionality)
    @PutMapping("/{id}")
    public ResponseEntity<Enquiry> updateEnquiry(
            @PathVariable Integer id,
            @RequestBody EnquiryUpdateRequestDTO dto) {
        Enquiry updatedEnquiry = enquiryService.updateEnquiry(id, dto);
        return ResponseEntity.ok(updatedEnquiry);
    }

    // Follow-ups for logged-in staff (TODAY + pending)
    @GetMapping("/followups/staff/{staffId}")
    public List<EnquiryResponseDTO> getUpcomingFollowups(@PathVariable Integer staffId) {
        return enquiryService.getUpcomingFollowupsForStaff(staffId);
    }

    // View All follow-ups (Admin)
    @GetMapping("/followups/all")
    public List<EnquiryResponseDTO> getAllFollowups() {
        return enquiryService.getAllPendingFollowups();
    }

    // CALL button action (Follow-up update)
    @PutMapping("/followup")
    public Enquiry updateFollowup(@RequestBody EnquiryFollowUpRequestDTO dto) {
        return enquiryService.updateFollowup(dto);
    }
}