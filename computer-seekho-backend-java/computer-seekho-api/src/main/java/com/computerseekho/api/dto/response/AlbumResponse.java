package com.computerseekho.api.dto.response;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class AlbumResponse {
    private Integer albumId;
    private String albumName;
    private String albumDescription;
    private List<ImageResponse> images;
}

