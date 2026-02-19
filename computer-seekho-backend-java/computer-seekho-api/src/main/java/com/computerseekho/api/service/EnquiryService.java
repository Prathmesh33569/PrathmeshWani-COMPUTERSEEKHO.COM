package com.computerseekho.api.service;

import com.computerseekho.api.dto.request.EnquiryCreateRequestDTO;
import com.computerseekho.api.dto.request.EnquiryFollowUpRequestDTO;
import com.computerseekho.api.dto.request.EnquiryUpdateRequestDTO;
import com.computerseekho.api.dto.response.EnquiryResponseDTO;
import com.computerseekho.api.entity.Enquiry;

import java.util.List;

public interface EnquiryService {

    Enquiry createEnquiry(EnquiryCreateRequestDTO dto);

    Enquiry getEnquiryById(Integer id);

    Enquiry updateEnquiry(Integer id, EnquiryUpdateRequestDTO dto);

    List<EnquiryResponseDTO> getUpcomingFollowupsForStaff(Integer staffId);

    List<EnquiryResponseDTO> getAllPendingFollowups();

    Enquiry updateFollowup(EnquiryFollowUpRequestDTO dto);
}