package com.computerseekho.api.service;

import com.computerseekho.api.dto.response.AlbumResponseDTO;
import com.computerseekho.api.dto.response.ImageResponseDTO;
import com.computerseekho.api.entity.ProgramCode;

import java.util.List;

public interface PlacementService {

    List<AlbumResponseDTO> getPlacementAlbums(ProgramCode programCode);

    List<ImageResponseDTO> getPlacementAlbumImages(Integer albumId);
}
