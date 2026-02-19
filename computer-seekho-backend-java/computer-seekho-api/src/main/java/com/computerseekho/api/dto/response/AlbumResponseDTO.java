package com.computerseekho.api.dto.response;

import java.time.LocalDateTime;

public record AlbumResponseDTO(
        Integer albumId,
        String albumName,
        String albumDescription,
        String coverImagePath,
        LocalDateTime startDate
) {}
