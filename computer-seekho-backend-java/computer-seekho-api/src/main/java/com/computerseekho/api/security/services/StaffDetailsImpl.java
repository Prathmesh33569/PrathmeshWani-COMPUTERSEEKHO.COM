package com.computerseekho.api.security.services;

import com.computerseekho.api.entity.Staff;
import com.fasterxml.jackson.annotation.JsonIgnore;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.UserDetails;

import java.util.Collection;
import java.util.Collections;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class StaffDetailsImpl implements UserDetails {

    private static final long serialVersionUID = 1L;

    private Integer staffId;
    private String staffUsername;
    private String staffEmail;
    private String staffName;
    private String staffRole;

    @JsonIgnore
    private String staffPassword;

    private Collection<? extends GrantedAuthority> authorities;

    // Build StaffDetailsImpl from Staff entity
    public static StaffDetailsImpl build(Staff staff) {
        // Convert staff role to Spring Security authority
        GrantedAuthority authority;
        if ("teaching".equalsIgnoreCase(staff.getStaffRole())) {
            authority = new SimpleGrantedAuthority("ROLE_TEACHING");
        } else {
            authority = new SimpleGrantedAuthority("ROLE_NON_TEACHING");
        }

        return new StaffDetailsImpl(
                staff.getStaffId(),
                staff.getStaffUsername(),
                staff.getStaffEmail(),
                staff.getStaffName(),
                staff.getStaffRole(),
                staff.getStaffPassword(),
                Collections.singletonList(authority)
        );
    }

    @Override
    public Collection<? extends GrantedAuthority> getAuthorities() {
        return authorities;
    }

    @Override
    public String getPassword() {
        return staffPassword;
    }

    @Override
    public String getUsername() {
        return staffUsername;
    }

    @Override
    public boolean isAccountNonExpired() {
        return true;
    }

    @Override
    public boolean isAccountNonLocked() {
        return true;
    }

    @Override
    public boolean isCredentialsNonExpired() {
        return true;
    }

    @Override
    public boolean isEnabled() {
        return true;
    }
}