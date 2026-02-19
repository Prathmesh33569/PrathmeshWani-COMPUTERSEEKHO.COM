package com.computerseekho.api.security.services;

import com.computerseekho.api.entity.Staff;
import com.computerseekho.api.repository.StaffRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

@Service
public class StaffDetailsServiceImpl implements UserDetailsService {

    @Autowired
    private StaffRepository staffRepository;

    @Override
    @Transactional
    public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
        Staff staff = staffRepository.findByStaffUsername(username)
                .orElseThrow(() -> new UsernameNotFoundException("Staff Not Found with username: " + username));

        return StaffDetailsImpl.build(staff);
    }
}