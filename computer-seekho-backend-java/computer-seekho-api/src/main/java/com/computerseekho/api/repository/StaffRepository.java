package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Staff;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface StaffRepository extends JpaRepository<Staff, Integer> {

    Optional<Staff> findByStaffUsername(String staffUsername);

    Optional<Staff> findByStaffEmail(String staffEmail);

    Boolean existsByStaffUsername(String staffUsername);

    Boolean existsByStaffEmail(String staffEmail);

    // 🆕 Get all teaching staff
    List<Staff> findByStaffRole(String staffRole);
}