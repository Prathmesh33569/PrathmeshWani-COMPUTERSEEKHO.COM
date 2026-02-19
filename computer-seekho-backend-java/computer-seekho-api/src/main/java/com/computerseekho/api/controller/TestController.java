package com.computerseekho.api.controller;

import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@CrossOrigin(origins = "*", maxAge = 3600)
@RestController
@RequestMapping("/api/test")
public class TestController {

    @GetMapping("/all")
    public String allAccess() {
        return "Public Content - Anyone can access this!";
    }

    @GetMapping("/teaching")
    @PreAuthorize("hasRole('TEACHING')")
    public String teachingAccess() {
        return "Teaching Staff Content - Only teaching staff can access this!";
    }

    @GetMapping("/non-teaching")

    @PreAuthorize("hasRole('NON_TEACHING') or hasRole('TEACHING')")
    public String nonTeachingAccess() {
        return "Non-Teaching Staff Content - access by non-teaching staff and teaching staff can access this!";

    }
}