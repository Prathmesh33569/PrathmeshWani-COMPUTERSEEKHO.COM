package com.computerseekho.api.service;

import com.computerseekho.api.entity.Staff;
import com.computerseekho.api.repository.StaffRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class StaffService {

    @Autowired
    private StaffRepository staffRepository;

    /**
     * Get all teaching staff
     */
    public List<Staff> getFacultyList() {
        return staffRepository.findByStaffRole("teaching");
    }

    /**
     * Get all staff
     */
    public List<Staff> getAllStaff() {
        return staffRepository.findAll();
    }

    /**
     * Get staff by ID
     */
    public Staff getStaffById(Integer staffId) {
        return staffRepository.findById(staffId)
                .orElseThrow(() -> new RuntimeException("Staff not found with id: " + staffId));
    }
}
