package com.computerseekho.api.service;

import com.computerseekho.api.repository.RecruitersRepository;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class RecruiterService {

    private final RecruitersRepository recruitersRepository;

    public RecruiterService(RecruitersRepository recruitersRepository) {
        this.recruitersRepository = recruitersRepository;
    }

    public List<String> getRecruiterImages() {
        return recruitersRepository.findRecruiterImages();
    }
}
