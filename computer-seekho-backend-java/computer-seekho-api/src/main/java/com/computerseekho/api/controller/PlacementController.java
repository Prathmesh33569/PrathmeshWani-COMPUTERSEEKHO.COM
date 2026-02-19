package com.computerseekho.api.controller;

import com.computerseekho.api.dto.response.AlbumResponseDTO;
import com.computerseekho.api.dto.response.ImageResponseDTO;
import com.computerseekho.api.entity.ProgramCode;
import com.computerseekho.api.service.PlacementService;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/placements")
@CrossOrigin(origins = "*")
public class PlacementController {

    private final PlacementService placementService;

    public PlacementController(PlacementService placementService) {
        this.placementService = placementService;
    }

    @GetMapping("/albums")
    public List<AlbumResponseDTO> getPlacementAlbums(
            @RequestParam ProgramCode program
    ) {
        return placementService.getPlacementAlbums(program);
    }

    @GetMapping("/albums/{albumId}/images")
    public List<ImageResponseDTO> getPlacementAlbumImages(
            @PathVariable Integer albumId
    ) {
        return placementService.getPlacementAlbumImages(albumId);
    }}
