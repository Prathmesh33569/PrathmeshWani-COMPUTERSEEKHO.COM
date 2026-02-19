package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Album;
import com.computerseekho.api.entity.AlbumType;
import com.computerseekho.api.entity.ProgramCode;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface PlacementRepository extends JpaRepository<Album,Integer> {
    List<Album> findByAlbumTypeAndProgramCodeAndAlbumIsActiveTrueOrderByStartDateDesc(
            AlbumType albumType,
            ProgramCode programCode
    );
}
