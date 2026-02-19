package com.computerseekho.api.controller;

import com.computerseekho.api.entity.Staff;
import com.computerseekho.api.service.StaffService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@CrossOrigin(origins = "*", maxAge = 3600)
@RestController
@RequestMapping("/api/staff")
public class StaffController {

    @Autowired
    private StaffService staffService;

    /**
     * Get all teaching staff (faculty)
     * GET /api/staff/faculty
     */
    @GetMapping("/faculty")
    public ResponseEntity<List<Staff>> getAllFaculty() {
        List<Staff> faculty = staffService.getFacultyList();
        return ResponseEntity.ok(faculty);
    }

    /**
     * Get all staff
     * GET /api/staff/all
     */
    @GetMapping("/all")
    public ResponseEntity<List<Staff>> getAllStaff() {
        List<Staff> staff = staffService.getAllStaff();
        return ResponseEntity.ok(staff);
    }

    /**
     * Get staff by ID
     * GET /api/staff/{id}
     */
    @GetMapping("/{id}")
    public ResponseEntity<Staff> getStaffById(@PathVariable Integer id) {
        Staff staff = staffService.getStaffById(id);
        return ResponseEntity.ok(staff);
    }
}
