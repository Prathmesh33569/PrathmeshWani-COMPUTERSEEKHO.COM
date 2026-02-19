package com.computerseekho.api.dto.response;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ImageResponse {
    private Integer imageId;
    private String imagePath;
    private Boolean isAlbumCover;
}

