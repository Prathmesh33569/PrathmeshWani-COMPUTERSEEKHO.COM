package com.computerseekho.api.service;

import com.computerseekho.api.dto.response.AlbumResponse;
import com.computerseekho.api.dto.response.ImageResponse;
import com.computerseekho.api.entity.Album;
import com.computerseekho.api.entity.AlbumType;
import com.computerseekho.api.repository.AlbumRepository;
import com.computerseekho.api.repository.ImageRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class AlbumService {

    @Autowired
    private AlbumRepository albumRepository;

    @Autowired
    private ImageRepository imageRepository;

    public List<AlbumResponse> getCampusLifeAlbums() {

        // 1️⃣ Get CAMPUS_LIFE albums
        List<Album> albums = albumRepository
                .findByAlbumTypeAndAlbumIsActive(
                        AlbumType.CAMPUS_LIFE,
                        true
                );

        // 2️⃣ Convert to response
        return albums.stream().map(album -> {

            // 3️⃣ Get images for album
            List<ImageResponse> images = imageRepository
                    .findByAlbumAlbumIdAndImageIsActive(
                            album.getAlbumId(),
                            true
                    )
                    .stream()
                    .map(img -> new ImageResponse(
                            img.getImageId(),
                            img.getImagePath(),
                            img.getIsAlbumCover()
                    ))
                    .toList();

            // 4️⃣ Create AlbumResponse
            AlbumResponse response = new AlbumResponse();
            response.setAlbumId(album.getAlbumId());
            response.setAlbumName(album.getAlbumName());
            response.setAlbumDescription(album.getAlbumDescription());
            response.setImages(images);

            return response;

        }).toList();
    }
}
