package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Image;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface PlacementImageRepository extends JpaRepository<Image, Integer> {

    Optional<Image> findByAlbum_AlbumIdAndIsAlbumCoverTrueAndImageIsActiveTrue(Integer albumId);

    List<Image> findByAlbum_AlbumIdAndImageIsActiveTrue(Integer albumId);
}
