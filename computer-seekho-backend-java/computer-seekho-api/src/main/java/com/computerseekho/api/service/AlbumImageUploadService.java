package com.computerseekho.api.service;

import com.computerseekho.api.entity.Image;
import com.computerseekho.api.repository.ImageRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

@Service
public class AlbumImageUploadService {

    @Autowired
    private ImageRepository imageRepository;

    @Transactional
    public void saveAll(List<Image> images) {
        imageRepository.saveAll(images);
    }
}
