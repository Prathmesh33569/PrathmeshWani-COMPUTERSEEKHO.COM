package com.computerseekho.api.repository;


import com.computerseekho.api.entity.Image;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ImageRepository extends JpaRepository<Image, Integer> {

    List<Image> findByAlbumAlbumIdAndImageIsActive(
            Integer albumId,
            Boolean imageIsActive
    );
}
