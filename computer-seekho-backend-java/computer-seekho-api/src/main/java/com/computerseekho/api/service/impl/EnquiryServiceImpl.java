package com.computerseekho.api.service.impl;

import com.computerseekho.api.dto.request.EnquiryCreateRequestDTO;
import com.computerseekho.api.dto.request.EnquiryFollowUpRequestDTO;
import com.computerseekho.api.dto.request.EnquiryUpdateRequestDTO;
import com.computerseekho.api.dto.response.EnquiryResponseDTO;
import com.computerseekho.api.entity.*;
import com.computerseekho.api.repository.*;
import com.computerseekho.api.service.EnquiryService;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDate;
import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class EnquiryServiceImpl implements EnquiryService {

    private final EnquiryRepository enquiryRepository;
    private final CourseRepository courseRepository;
    private final StaffRepository staffRepository;
    private final ClosureReasonRepository closureReasonRepository;

    @Override
    @Transactional
    public Enquiry createEnquiry(EnquiryCreateRequestDTO dto) {

        Enquiry enquiry = new Enquiry();

        enquiry.setEnquirerName(dto.getEnquirerName());
        enquiry.setStudentName(dto.getStudentName());
        enquiry.setEnquirerAddress(dto.getEnquirerAddress());
        enquiry.setEnquirerMobile(dto.getEnquirerMobile());
        enquiry.setEnquirerAlternateMobile(dto.getEnquirerAlternateMobile());
        enquiry.setEnquirerEmailId(dto.getEnquirerEmailId());
        enquiry.setEnquirerQuery(dto.getEnquirerQuery());
        enquiry.setEnquirySource(dto.getEnquirySource());

        enquiry.setCourse(courseRepository.findById(dto.getCourseId())
                .orElseThrow(() -> new RuntimeException("Course not found with id: " + dto.getCourseId())));

        enquiry.setStaff(staffRepository.findById(dto.getStaffId())
                .orElseThrow(() -> new RuntimeException("Staff not found with id: " + dto.getStaffId())));

        enquiry.setFollowupDate(
                dto.getFollowupDate() != null
                        ? dto.getFollowupDate()
                        : LocalDate.now().plusDays(3)
        );

        return enquiryRepository.save(enquiry);
    }

    @Override
    public List<EnquiryResponseDTO> getUpcomingFollowupsForStaff(Integer staffId) {

        Staff staff = staffRepository.findById(staffId)
                .orElseThrow(() -> new RuntimeException("Staff not found with id: " + staffId));

        return enquiryRepository
                .findByStaffAndIsClosedFalseAndFollowupDateGreaterThanEqual(
                        staff, LocalDate.now()
                )
                .stream()
                .map(this::mapToResponseDTO)
                .collect(Collectors.toList());
    }

    @Override
    public List<EnquiryResponseDTO> getAllPendingFollowups() {

        return enquiryRepository
                .findByIsClosedFalseAndFollowupDateGreaterThanEqual(LocalDate.now())
                .stream()
                .map(this::mapToResponseDTO)
                .collect(Collectors.toList());
    }

    @Override
    public Enquiry getEnquiryById(Integer id) {
        return enquiryRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Enquiry not found with id: " + id));
    }

    @Override
    @Transactional
    public Enquiry updateEnquiry(Integer id, EnquiryUpdateRequestDTO dto) {
        Enquiry enquiry = enquiryRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Enquiry not found with id: " + id));

        // Update basic details
        if (dto.getEnquirerName() != null && !dto.getEnquirerName().trim().isEmpty()) {
            enquiry.setEnquirerName(dto.getEnquirerName());
        }

        if (dto.getStudentName() != null) {
            enquiry.setStudentName(dto.getStudentName());
        }

        if (dto.getEnquirerAddress() != null) {
            enquiry.setEnquirerAddress(dto.getEnquirerAddress());
        }

        if (dto.getEnquirerMobile() != null) {
            enquiry.setEnquirerMobile(dto.getEnquirerMobile());
        }

        if (dto.getEnquirerAlternateMobile() != null) {
            enquiry.setEnquirerAlternateMobile(dto.getEnquirerAlternateMobile());
        }

        if (dto.getEnquirerEmailId() != null) {
            enquiry.setEnquirerEmailId(dto.getEnquirerEmailId());
        }

        if (dto.getEnquirerQuery() != null) {
            enquiry.setEnquirerQuery(dto.getEnquirerQuery());
        }

        if (dto.getEnquirySource() != null) {
            enquiry.setEnquirySource(dto.getEnquirySource());
        }

        // Update course if changed
        if (dto.getCourseId() != null) {
            Course course = courseRepository.findById(dto.getCourseId())
                    .orElseThrow(() -> new RuntimeException("Course not found with id: " + dto.getCourseId()));
            enquiry.setCourse(course);
        }

        // Update follow-up date if provided
        if (dto.getFollowupDate() != null) {
            enquiry.setFollowupDate(dto.getFollowupDate());
        }

        // Update special instructions
        if (dto.getSpecialInstructions() != null) {
            enquiry.setSpecialInstructions(dto.getSpecialInstructions());
        }

        return enquiryRepository.save(enquiry);
    }

    @Override
    @Transactional
    public Enquiry updateFollowup(EnquiryFollowUpRequestDTO dto) {

        Enquiry enquiry = enquiryRepository.findById(dto.getEnquiryId())
                .orElseThrow(() -> new RuntimeException("Enquiry not found with id: " + dto.getEnquiryId()));

        // Update remarks and special instructions
        if (dto.getSpecialInstructions() != null && !dto.getSpecialInstructions().trim().isEmpty()) {
            enquiry.setSpecialInstructions(dto.getSpecialInstructions());
        }

        // Increment inquiry counter
        enquiry.setInquiryCounter(enquiry.getInquiryCounter() + 1);

        // Handle closure or next follow-up
        if (dto.getCloseEnquiry() != null && dto.getCloseEnquiry()) {
            // Close the enquiry
            enquiry.setIsClosed(true);

            // Set closure reason if provided
            if (dto.getClosureReasonId() != null) {
                ClosureReason closureReason = closureReasonRepository.findById(dto.getClosureReasonId())
                        .orElseThrow(() -> new RuntimeException("Closure reason not found with id: " + dto.getClosureReasonId()));
                enquiry.setClosureReason(closureReason);
            }

            // Set closure reason text (for "Other" option)
            if (dto.getClosureReasonText() != null && !dto.getClosureReasonText().trim().isEmpty()) {
                enquiry.setClosureReasonText(dto.getClosureReasonText());
            }

        } else {
            // Set next follow-up date
            enquiry.setFollowupDate(
                    dto.getNextFollowupDate() != null
                            ? dto.getNextFollowupDate()
                            : LocalDate.now().plusDays(3)
            );
        }

        return enquiryRepository.save(enquiry);
    }

    private EnquiryResponseDTO mapToResponseDTO(Enquiry enquiry) {
        EnquiryResponseDTO dto = new EnquiryResponseDTO();
        dto.setEnquiryId(enquiry.getEnquiryId());
        dto.setEnquirerName(enquiry.getEnquirerName());
        dto.setEnquirerMobile(enquiry.getEnquirerMobile());
        dto.setCourseName(enquiry.getCourse() != null ? enquiry.getCourse().getCourseName() : "N/A");
        dto.setFollowupDate(enquiry.getFollowupDate());
        dto.setIsClosed(enquiry.getIsClosed());
        return dto;
    }
}