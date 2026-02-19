package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Enquiry;
import com.computerseekho.api.entity.Staff;
import org.springframework.data.jpa.repository.JpaRepository;

import java.time.LocalDate;
import java.util.List;

public interface EnquiryRepository extends JpaRepository<Enquiry, Integer> {

    // Follow-ups from today onwards
    List<Enquiry> findByStaffAndIsClosedFalseAndFollowupDateGreaterThanEqual(
            Staff staff,
            LocalDate date
    );


    // All pending follow-ups (admin / view all)
    List<Enquiry> findByIsClosedFalseAndFollowupDateGreaterThanEqual(LocalDate date);
}
