package com.computerseekho.api.service.impl;

import com.computerseekho.api.dto.response.AlbumResponseDTO;
import com.computerseekho.api.dto.response.ImageResponseDTO;
import com.computerseekho.api.entity.Album;
import com.computerseekho.api.entity.AlbumType;
import com.computerseekho.api.entity.Image;
import com.computerseekho.api.entity.ProgramCode;
import com.computerseekho.api.repository.PlacementImageRepository;
import com.computerseekho.api.repository.PlacementRepository;
import com.computerseekho.api.service.PlacementService;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class PlacementServiceImpl implements PlacementService {

    private final PlacementRepository placementRepository;
    private final PlacementImageRepository placementImageRepository;

    public PlacementServiceImpl(
            PlacementRepository placementRepository,
            PlacementImageRepository placementImageRepository
    ) {
        this.placementRepository = placementRepository;
        this.placementImageRepository = placementImageRepository;
    }

    @Override
    public
    List<AlbumResponseDTO> getPlacementAlbums(ProgramCode programCode) {

        List<Album> albums =
                placementRepository.findByAlbumTypeAndProgramCodeAndAlbumIsActiveTrueOrderByStartDateDesc(
                        AlbumType.PLACEMENT,
                        programCode
                );

        return albums.stream().map(album -> {

            String coverImage = placementImageRepository
                    .findByAlbum_AlbumIdAndIsAlbumCoverTrueAndImageIsActiveTrue(album.getAlbumId())
                    .map(Image::getImagePath)
                    .orElse(null);

            return new AlbumResponseDTO(
                    album.getAlbumId(),
                    album.getAlbumName(),
                    album.getAlbumDescription(),
                    coverImage,
                    album.getStartDate()
            );
        }).toList();
    }

    @Override
    public List<ImageResponseDTO> getPlacementAlbumImages(Integer albumId) {
        return placementImageRepository
                .findByAlbum_AlbumIdAndImageIsActiveTrue(albumId)
                .stream()
                .filter(img -> !img.getIsAlbumCover()) //  Exclude cover images
                .map(img -> new ImageResponseDTO(
                        img.getImageId(),
                        img.getImagePath()
                ))
                .toList();
    }
}

