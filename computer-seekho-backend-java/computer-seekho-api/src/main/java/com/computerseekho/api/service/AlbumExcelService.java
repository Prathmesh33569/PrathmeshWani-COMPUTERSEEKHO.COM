package com.computerseekho.api.service;

import com.computerseekho.api.entity.Album;
import com.computerseekho.api.entity.AlbumType;
import com.computerseekho.api.entity.Image;
import com.computerseekho.api.entity.ProgramCode;
import com.computerseekho.api.repository.AlbumRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.*;

@Service
public class AlbumExcelService {

    @Autowired
    private AlbumRepository albumRepository;

    public List<Image> parseExcel(MultipartFile file) throws Exception {

        List<Image> images = new ArrayList<>();

        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();

            // Skip header row
            if (rows.hasNext()) rows.next();

            while (rows.hasNext()) {

                Row row = rows.next();
                if (row == null || row.getCell(0) == null) continue;

                String albumTypeStr = getString(row.getCell(0));
                if (albumTypeStr == null || albumTypeStr.equalsIgnoreCase("album_type")) {
                    continue;
                }

                AlbumType albumType = AlbumType.valueOf(albumTypeStr.trim());

                // -------- Program Code (nullable) --------
                ProgramCode programCode = null;
                String programStr = getString(row.getCell(1));
                if (programStr != null && !programStr.isBlank()) {
                    programCode = ProgramCode.valueOf(programStr.trim());
                }

                final ProgramCode finalProgramCode = programCode;

                String albumName = getString(row.getCell(2));
                if (albumName == null) continue;

                Album album;

                // ===== CASE 1: programCode == NULL =====
                if (finalProgramCode == null) {

                    Optional<Album> existingAlbum =
                            albumRepository.findByAlbumTypeAndProgramCodeIsNull(albumType);

                    album = existingAlbum.orElseGet(() -> {
                        Album a = new Album();
                        a.setAlbumType(albumType);
                        a.setProgramCode(null);
                        a.setAlbumName(albumName);
                        a.setAlbumDescription(getString(row.getCell(3)));
                        a.setAlbumIsActive(getBoolean(row.getCell(4)));
                        return albumRepository.save(a);
                    });

                }
                // ===== CASE 2: programCode != NULL =====
                else {

                    album = albumRepository
                            .findByAlbumTypeAndProgramCodeAndAlbumName(
                                    albumType, finalProgramCode, albumName
                            )
                            .orElseGet(() -> {
                                Album a = new Album();
                                a.setAlbumType(albumType);
                                a.setProgramCode(finalProgramCode);
                                a.setAlbumName(albumName);
                                a.setAlbumDescription(getString(row.getCell(3)));
                                a.setAlbumIsActive(getBoolean(row.getCell(4)));
                                return albumRepository.save(a);
                            });
                }

                // -------- Image --------
                Image image = new Image();
                image.setAlbum(album);
                image.setImagePath(getString(row.getCell(5)));
                image.setIsAlbumCover(getBoolean(row.getCell(6)));
                image.setImageIsActive(getBoolean(row.getCell(7)));

                images.add(image);
            }
        }

        return images;
    }

    // ================= HELPER METHODS =================

    private String getString(Cell cell) {
        if (cell == null) return null;

        DataFormatter formatter = new DataFormatter();
        String value = formatter.formatCellValue(cell);

        return value != null ? value.trim() : null;
    }

    private boolean getBoolean(Cell cell) {
        if (cell == null) return false;

        DataFormatter formatter = new DataFormatter();
        String value = formatter.formatCellValue(cell).trim();

        return value.equalsIgnoreCase("true")
                || value.equalsIgnoreCase("yes")
                || value.equals("1");
    }
}
