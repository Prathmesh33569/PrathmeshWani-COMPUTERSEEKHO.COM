//package com.computerseekho.api.config;
//
//import com.computerseekho.api.security.jwt.AuthEntryPointJwt;
//import com.computerseekho.api.security.jwt.JwtAuthenticationFilter;
//import com.computerseekho.api.security.oauth2.OAuth2AuthenticationSuccessHandler;
//import com.computerseekho.api.security.services.StaffDetailsServiceImpl;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.context.annotation.Bean;
//import org.springframework.context.annotation.Configuration;
//import org.springframework.security.authentication.AuthenticationManager;
//import org.springframework.security.authentication.dao.DaoAuthenticationProvider;
//import org.springframework.security.config.annotation.authentication.configuration.AuthenticationConfiguration;
//import org.springframework.security.config.annotation.method.configuration.EnableMethodSecurity;
//import org.springframework.security.config.annotation.web.builders.HttpSecurity;
//import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
//import org.springframework.security.config.http.SessionCreationPolicy;
//import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
//import org.springframework.security.crypto.password.PasswordEncoder;
//import org.springframework.security.web.SecurityFilterChain;
//import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;
//
//
//import org.springframework.context.annotation.Bean;
//import org.springframework.context.annotation.Configuration;
//import org.springframework.security.config.annotation.web.builders.HttpSecurity;
//import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
//import org.springframework.security.web.SecurityFilterChain;
//import static org.springframework.security.config.Customizer.withDefaults;
//@Configuration
//@EnableWebSecurity
//@EnableMethodSecurity
//public class SecurityConfig {
//
//    @Autowired
//    private StaffDetailsServiceImpl staffDetailsService;
//
//    @Autowired
//    private AuthEntryPointJwt unauthorizedHandler;
//
//    @Autowired
//    private OAuth2AuthenticationSuccessHandler oAuth2AuthenticationSuccessHandler;
//
//    @Bean
//    public JwtAuthenticationFilter authenticationJwtTokenFilter() {
//        return new JwtAuthenticationFilter();
//    }
//
//    @Bean
//    public DaoAuthenticationProvider authenticationProvider() {
//        DaoAuthenticationProvider authProvider = new DaoAuthenticationProvider();
//
//        authProvider.setUserDetailsService(staffDetailsService);
//        authProvider.setPasswordEncoder(passwordEncoder());
//
//        return authProvider;
//    }
//
//    @Bean
//    public AuthenticationManager authenticationManager(AuthenticationConfiguration authConfig) throws Exception {
//        return authConfig.getAuthenticationManager();
//    }
//
//    @Bean
//    public PasswordEncoder passwordEncoder() {
//        return new BCryptPasswordEncoder();
//    }
//
//    @Bean
//    public SecurityFilterChain filterChain(HttpSecurity http) throws Exception {
//        http.cors(withDefaults())
//                .csrf(csrf -> csrf.disable())
//                .exceptionHandling(exception -> exception.authenticationEntryPoint(unauthorizedHandler))
//                .sessionManagement(session -> session.sessionCreationPolicy(SessionCreationPolicy.STATELESS))
//                .authorizeHttpRequests(auth ->
//                        auth.requestMatchers("/api/auth/**").permitAll()
//                                .requestMatchers("/oauth2/**").permitAll()
//                                .requestMatchers("/login/oauth2/**").permitAll()
//                                .requestMatchers("/api/test/**").permitAll()
//                                .requestMatchers("/api/announcements/**").permitAll()
//                                .requestMatchers("/api/enquiries/**").permitAll()
//                                .requestMatchers("/api/albums/**").permitAll()
//                                .requestMatchers("/api/courses/**").permitAll()
//                                .requestMatchers("/api/recruiters/**").permitAll()
//                                .requestMatchers("/api/placements/**").permitAll()
//                                .requestMatchers("/api/batches/**").permitAll()
//                                .requestMatchers("/api/students/**").permitAll()
//                                .requestMatchers("/api/payments/**").permitAll()
//                                .requestMatchers("/api/email/**").permitAll()
//                                .requestMatchers("/api/placements/**").permitAll()
//                                .requestMatchers("/api/staff/**").permitAll()
//                                .anyRequest().authenticated()
//                )
//                .oauth2Login(oauth2 -> oauth2
//                        .successHandler(oAuth2AuthenticationSuccessHandler)
//                );
//        http.authenticationProvider(authenticationProvider());
//        http.addFilterBefore(authenticationJwtTokenFilter(), UsernamePasswordAuthenticationFilter.class);
//
//        return http.build();
//    }
//}

package com.computerseekho.api.config;

import com.computerseekho.api.security.jwt.AuthEntryPointJwt;
import com.computerseekho.api.security.jwt.JwtAuthenticationFilter;
import com.computerseekho.api.security.oauth2.OAuth2AuthenticationSuccessHandler;
import com.computerseekho.api.security.services.StaffDetailsServiceImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.dao.DaoAuthenticationProvider;
import org.springframework.security.config.annotation.authentication.configuration.AuthenticationConfiguration;
import org.springframework.security.config.annotation.method.configuration.EnableMethodSecurity;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;


import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.web.SecurityFilterChain;
import static org.springframework.security.config.Customizer.withDefaults;
@Configuration
@EnableWebSecurity
@EnableMethodSecurity
public class SecurityConfig {

    @Autowired
    private StaffDetailsServiceImpl staffDetailsService;

    @Autowired
    private AuthEntryPointJwt unauthorizedHandler;

    @Autowired
    private OAuth2AuthenticationSuccessHandler oAuth2AuthenticationSuccessHandler;

    @Bean
    public JwtAuthenticationFilter authenticationJwtTokenFilter() {
        return new JwtAuthenticationFilter();
    }

    @Bean
    public DaoAuthenticationProvider authenticationProvider() {
        DaoAuthenticationProvider authProvider = new DaoAuthenticationProvider();

        authProvider.setUserDetailsService(staffDetailsService);
        authProvider.setPasswordEncoder(passwordEncoder());

        return authProvider;
    }

    @Bean
    public AuthenticationManager authenticationManager(AuthenticationConfiguration authConfig) throws Exception {
        return authConfig.getAuthenticationManager();
    }

    @Bean
    public PasswordEncoder passwordEncoder() {
        return new BCryptPasswordEncoder();
    }

    @Bean
    public SecurityFilterChain filterChain(HttpSecurity http) throws Exception {
        http.cors(withDefaults()) // Enable CORS
                .csrf(csrf -> csrf.disable())
                .exceptionHandling(exception -> exception.authenticationEntryPoint(unauthorizedHandler))
                .sessionManagement(session -> session.sessionCreationPolicy(SessionCreationPolicy.STATELESS))
                .authorizeHttpRequests(auth ->
                        auth.requestMatchers("/api/auth/**").permitAll()
                                .requestMatchers("/oauth2/**").permitAll()
                                .requestMatchers("/login/oauth2/**").permitAll()
                                .requestMatchers("/api/test/**").permitAll()
                                .requestMatchers("/api/announcements/**").permitAll()
                                .requestMatchers("/api/enquiries/**").permitAll()
                                .requestMatchers("/api/albums/**").permitAll()
                                .requestMatchers("/api/courses/**").permitAll()
                                .requestMatchers("/api/recruiters/**").permitAll()
                                .requestMatchers("/api/placements/**").permitAll()
                                .requestMatchers("/api/batches/**").permitAll()
                                .requestMatchers("/api/students/**").permitAll()
                                .requestMatchers("/api/payments/**").permitAll()
                                .requestMatchers("/api/email/**").permitAll()
                                .requestMatchers("/api/placements/**").permitAll()
                                .requestMatchers("/api/staff/**").permitAll()
                                .requestMatchers("/api/upload/**").permitAll() // Allow file upload
                                .requestMatchers("/api/receipt/**").permitAll()
                                .requestMatchers("/api/staff/faculty").permitAll()
                                .requestMatchers("/api/staff/all").permitAll()
                                .requestMatchers("/assets/**").permitAll()
                                .requestMatchers("/api/public/**").permitAll()
                                .requestMatchers("/actuator/**").permitAll()
                                .anyRequest().authenticated()
                )
                .oauth2Login(oauth2 -> oauth2
                        .successHandler(oAuth2AuthenticationSuccessHandler)
                );
        http.authenticationProvider(authenticationProvider());
        http.addFilterBefore(authenticationJwtTokenFilter(), UsernamePasswordAuthenticationFilter.class);

        return http.build();
    }
}