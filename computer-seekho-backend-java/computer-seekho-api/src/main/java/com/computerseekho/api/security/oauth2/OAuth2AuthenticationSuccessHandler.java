package com.computerseekho.api.security.oauth2;

import com.computerseekho.api.entity.Staff;
import com.computerseekho.api.repository.StaffRepository;
import com.computerseekho.api.security.jwt.JwtUtils;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.Authentication;
import org.springframework.security.oauth2.core.user.OAuth2User;
import org.springframework.security.web.authentication.SimpleUrlAuthenticationSuccessHandler;
import org.springframework.stereotype.Component;
import org.springframework.web.util.UriComponentsBuilder;

import java.io.IOException;
import java.util.Optional;

@Component
public class OAuth2AuthenticationSuccessHandler extends SimpleUrlAuthenticationSuccessHandler {

    @Autowired
    private JwtUtils jwtUtils;

    @Autowired
    private StaffRepository staffRepository;

    @Override
    public void onAuthenticationSuccess(HttpServletRequest request, HttpServletResponse response,
                                        Authentication authentication) throws IOException, ServletException {

        OAuth2User oAuth2User = (OAuth2User) authentication.getPrincipal();

        String email = oAuth2User.getAttribute("email");
        String name = oAuth2User.getAttribute("name");

        // Find staff by email
        Optional<Staff> staffOptional = staffRepository.findByStaffEmail(email);

        if (staffOptional.isEmpty()) {
            // Redirect to login with error
            String targetUrl = UriComponentsBuilder.fromUriString("http://localhost:5173/admin/login")
                    .queryParam("error", "true")
                    .queryParam("message", "No staff account found with email: " + email)
                    .build().toUriString();

            getRedirectStrategy().sendRedirect(request, response, targetUrl);
            return;
        }

        Staff staff = staffOptional.get();

        // Generate JWT token
        String jwt = jwtUtils.generateTokenFromUsername(staff.getStaffUsername());

        // Redirect to /oauth-success with token
        String targetUrl = UriComponentsBuilder.fromUriString("http://localhost:5173/oauth-success")
                .queryParam("token", jwt)
                .queryParam("email", email)
                .queryParam("name", staff.getStaffName())
                .build().toUriString();

        getRedirectStrategy().sendRedirect(request, response, targetUrl);
    }
}