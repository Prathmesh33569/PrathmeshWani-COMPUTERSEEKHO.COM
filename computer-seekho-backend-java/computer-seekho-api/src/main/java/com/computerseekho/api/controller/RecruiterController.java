package com.computerseekho.api.controller;

import com.computerseekho.api.service.RecruiterService;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/recruiters")
@CrossOrigin(origins = "*")
public class RecruiterController {

    private final RecruiterService recruiterService;

    public RecruiterController(RecruiterService recruiterService) {
        this.recruiterService = recruiterService;
    }

    @GetMapping("/images")
    public List<String> getRecruiterImages() {
        return recruiterService.getRecruiterImages();
    }
}
