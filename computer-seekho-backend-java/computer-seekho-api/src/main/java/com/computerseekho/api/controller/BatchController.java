//package com.computerseekho.api.controller;
//
//import com.computerseekho.api.dto.request.BatchCreateRequestDTO;
//import com.computerseekho.api.dto.request.BatchUpdateRequestDTO;
//import com.computerseekho.api.dto.response.BatchResponseDTO;
//import com.computerseekho.api.entity.Batch;
//import com.computerseekho.api.entity.Course;
//import com.computerseekho.api.service.BatchService;
//import com.computerseekho.api.service.CourseService;
//import org.springframework.http.HttpStatus;
//import org.springframework.http.ResponseEntity;
//import org.springframework.web.bind.annotation.*;
//
//import java.util.List;
//import java.util.stream.Collectors;
//
//@RestController
//@RequestMapping("/api/batches")
//@CrossOrigin(origins = "*")
//public class BatchController {
//
//    private final BatchService batchService;
//    private final CourseService courseService;
//
//    public BatchController(BatchService batchService, CourseService courseService) {
//        this.batchService = batchService;
//        this.courseService = courseService;
//    }
//
//    // CREATE
//    @PostMapping
//    public ResponseEntity<BatchResponseDTO> createBatch(@RequestBody BatchCreateRequestDTO requestDTO) {
//        try {
//            Batch batch = convertCreateRequestToEntity(requestDTO);
//            Batch savedBatch = batchService.createBatch(batch);
//            return new ResponseEntity<>(convertToResponseDTO(savedBatch), HttpStatus.CREATED);
//        } catch (Exception e) {
//            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }
//
//    // READ ALL
//    @GetMapping
//    public ResponseEntity<List<BatchResponseDTO>> getAllBatches() {
//        try {
//            List<Batch> batches = batchService.getAllBatches();
//            List<BatchResponseDTO> responseDTOs = batches.stream()
//                    .map(this::convertToResponseDTO)
//                    .collect(Collectors.toList());
//            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
//        } catch (Exception e) {
//            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }
//
//    // READ BY ID
//    @GetMapping("/{id}")
//    public ResponseEntity<BatchResponseDTO> getBatchById(@PathVariable Integer id) {
//        try {
//            Batch batch = batchService.getBatchById(id);
//            return new ResponseEntity<>(convertToResponseDTO(batch), HttpStatus.OK);
//        } catch (RuntimeException e) {
//            return new ResponseEntity<>(null, HttpStatus.NOT_FOUND);
//        }
//    }
//
//    // UPDATE
//    @PutMapping("/{id}")
//    public ResponseEntity<BatchResponseDTO> updateBatch(
//            @PathVariable Integer id,
//            @RequestBody BatchUpdateRequestDTO requestDTO) {
//        try {
//            Batch batch = convertUpdateRequestToEntity(requestDTO);
//            Batch updatedBatch = batchService.updateBatch(id, batch);
//            return new ResponseEntity<>(convertToResponseDTO(updatedBatch), HttpStatus.OK);
//        } catch (RuntimeException e) {
//            return new ResponseEntity<>(null, HttpStatus.NOT_FOUND);
//        }
//    }
//
//    // DELETE
//    @DeleteMapping("/{id}")
//    public ResponseEntity<HttpStatus> deleteBatch(@PathVariable Integer id) {
//        try {
//            batchService.deleteBatch(id);
//            return new ResponseEntity<>(HttpStatus.NO_CONTENT);
//        } catch (Exception e) {
//            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }
//
//    // SEARCH
//    @GetMapping("/search")
//    public ResponseEntity<List<BatchResponseDTO>> searchBatches(@RequestParam String keyword) {
//        try {
//            List<Batch> batches = batchService.searchBatch(keyword);
//            List<BatchResponseDTO> responseDTOs = batches.stream()
//                    .map(this::convertToResponseDTO)
//                    .collect(Collectors.toList());
//            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
//        } catch (Exception e) {
//            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }
//
//    // GET ACTIVE BATCHES ONLY
//    @GetMapping("/active")
//    public ResponseEntity<List<BatchResponseDTO>> getActiveBatches() {
//        try {
//            List<Batch> batches = batchService.getAllBatches().stream()
//                    .filter(Batch::getBatchIsActive)
//                    .collect(Collectors.toList());
//            List<BatchResponseDTO> responseDTOs = batches.stream()
//                    .map(this::convertToResponseDTO)
//                    .collect(Collectors.toList());
//            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
//        } catch (Exception e) {
//            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }
//
//    // GET BATCHES BY COURSE ID
//    @GetMapping("/course/{courseId}")
//    public ResponseEntity<List<BatchResponseDTO>> getBatchesByCourse(@PathVariable Integer courseId) {
//        try {
//            List<Batch> batches = batchService.getAllBatches().stream()
//                    .filter(batch -> batch.getCourse() != null &&
//                            batch.getCourse().getCourseId().equals(courseId))
//                    .filter(Batch::getBatchIsActive)
//                    .collect(Collectors.toList());
//
//            List<BatchResponseDTO> responseDTOs = batches.stream()
//                    .map(this::convertToResponseDTO)
//                    .collect(Collectors.toList());
//
//            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
//        } catch (Exception e) {
//            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }
//
//    // ===========================
//    // HELPER METHODS - DTO CONVERSION
//    // ===========================
//
//    private BatchResponseDTO convertToResponseDTO(Batch batch) {
//        BatchResponseDTO dto = new BatchResponseDTO();
//        dto.setBatchId(batch.getBatchId());
//        dto.setBatchName(batch.getBatchName());
//        dto.setBatchStartDate(batch.getBatchStartDate());
//        dto.setBatchEndDate(batch.getBatchEndDate());
//        dto.setCourseId(batch.getCourse() != null ? batch.getCourse().getCourseId() : null);
//        dto.setCourseName(batch.getCourse() != null ? batch.getCourse().getCourseName() : null);
//        dto.setPresentationDate(batch.getPresentationDate());
//        dto.setCourseFees(batch.getCourseFees());
//        dto.setCourseFeesFrom(batch.getCourseFeesFrom());
//        dto.setCourseFeesTo(batch.getCourseFeesTo());
//        dto.setBatchIsActive(batch.getBatchIsActive());
//        dto.setFinalPresentationDate(batch.getFinalPresentationDate());
//        dto.setBatchLogoUrl(batch.getBatchLogoUrl());
//        dto.setCreatedAt(batch.getCreatedAt());
//        dto.setUpdatedAt(batch.getUpdatedAt());
//        return dto;
//    }
//
//    private Batch convertCreateRequestToEntity(BatchCreateRequestDTO dto) {
//        Batch batch = new Batch();
//        batch.setBatchName(dto.getBatchName());
//        batch.setBatchStartDate(dto.getBatchStartDate());
//        batch.setBatchEndDate(dto.getBatchEndDate());
//
//        if (dto.getCourseId() != null) {
//            Course course = courseService.getCourseEntityById(dto.getCourseId());
//            batch.setCourse(course);
//        }
//
//        batch.setPresentationDate(dto.getPresentationDate());
//        batch.setCourseFees(dto.getCourseFees());
//        batch.setCourseFeesFrom(dto.getCourseFeesFrom());
//        batch.setCourseFeesTo(dto.getCourseFeesTo());
//        batch.setBatchIsActive(dto.getBatchIsActive());
//        batch.setFinalPresentationDate(dto.getFinalPresentationDate());
//        batch.setBatchLogoUrl(dto.getBatchLogoUrl());
//        return batch;
//    }
//
//    private Batch convertUpdateRequestToEntity(BatchUpdateRequestDTO dto) {
//        Batch batch = new Batch();
//        batch.setBatchName(dto.getBatchName());
//        batch.setBatchStartDate(dto.getBatchStartDate());
//        batch.setBatchEndDate(dto.getBatchEndDate());
//
//        if (dto.getCourseId() != null) {
//            Course course = courseService.getCourseEntityById(dto.getCourseId());
//            batch.setCourse(course);
//        }
//
//        batch.setPresentationDate(dto.getPresentationDate());
//        batch.setCourseFees(dto.getCourseFees());
//        batch.setCourseFeesFrom(dto.getCourseFeesFrom());
//        batch.setCourseFeesTo(dto.getCourseFeesTo());
//        batch.setBatchIsActive(dto.getBatchIsActive());
//        batch.setFinalPresentationDate(dto.getFinalPresentationDate());
//        batch.setBatchLogoUrl(dto.getBatchLogoUrl());
//        return batch;
//    }
//}

package com.computerseekho.api.controller;

import com.computerseekho.api.dto.request.BatchCreateRequestDTO;
import com.computerseekho.api.dto.request.BatchUpdateRequestDTO;
import com.computerseekho.api.dto.response.BatchResponseDTO;
import com.computerseekho.api.entity.Batch;
import com.computerseekho.api.entity.Course;
import com.computerseekho.api.service.BatchService;
import com.computerseekho.api.service.CourseService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/api/batches")
@CrossOrigin(origins = "*")
public class BatchController {

    private final BatchService batchService;
    private final CourseService courseService;

    public BatchController(BatchService batchService, CourseService courseService) {
        this.batchService = batchService;
        this.courseService = courseService;
    }

    // CREATE
    @PostMapping
    public ResponseEntity<BatchResponseDTO> createBatch(@RequestBody BatchCreateRequestDTO requestDTO) {
        try {
            Batch batch = convertCreateRequestToEntity(requestDTO);
            Batch savedBatch = batchService.createBatch(batch);
            return new ResponseEntity<>(convertToResponseDTO(savedBatch), HttpStatus.CREATED);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // READ ALL
    @GetMapping
    public ResponseEntity<List<BatchResponseDTO>> getAllBatches() {
        try {
            List<Batch> batches = batchService.getAllBatches();
            List<BatchResponseDTO> responseDTOs = batches.stream()
                    .map(this::convertToResponseDTO)
                    .collect(Collectors.toList());
            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // READ BY ID
    @GetMapping("/{id}")
    public ResponseEntity<BatchResponseDTO> getBatchById(@PathVariable Integer id) {
        try {
            Batch batch = batchService.getBatchById(id);
            return new ResponseEntity<>(convertToResponseDTO(batch), HttpStatus.OK);
        } catch (RuntimeException e) {
            return new ResponseEntity<>(null, HttpStatus.NOT_FOUND);
        }
    }

    // UPDATE
    @PutMapping("/{id}")
    public ResponseEntity<BatchResponseDTO> updateBatch(
            @PathVariable Integer id,
            @RequestBody BatchUpdateRequestDTO requestDTO) {
        try {
            Batch batch = convertUpdateRequestToEntity(requestDTO);
            Batch updatedBatch = batchService.updateBatch(id, batch);
            return new ResponseEntity<>(convertToResponseDTO(updatedBatch), HttpStatus.OK);
        } catch (RuntimeException e) {
            return new ResponseEntity<>(null, HttpStatus.NOT_FOUND);
        }
    }

    // DELETE
    @DeleteMapping("/{id}")
    public ResponseEntity<HttpStatus> deleteBatch(@PathVariable Integer id) {
        try {
            batchService.deleteBatch(id);
            return new ResponseEntity<>(HttpStatus.NO_CONTENT);
        } catch (Exception e) {
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // SEARCH
    @GetMapping("/search")
    public ResponseEntity<List<BatchResponseDTO>> searchBatches(@RequestParam String keyword) {
        try {
            List<Batch> batches = batchService.searchBatch(keyword);
            List<BatchResponseDTO> responseDTOs = batches.stream()
                    .map(this::convertToResponseDTO)
                    .collect(Collectors.toList());
            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // GET ACTIVE BATCHES ONLY
    @GetMapping("/active")
    public ResponseEntity<List<BatchResponseDTO>> getActiveBatches() {
        try {
            List<Batch> batches = batchService.getAllBatches().stream()
                    .filter(Batch::getBatchIsActive)
                    .collect(Collectors.toList());
            List<BatchResponseDTO> responseDTOs = batches.stream()
                    .map(this::convertToResponseDTO)
                    .collect(Collectors.toList());
            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // GET BATCHES BY COURSE ID
    @GetMapping("/course/{courseId}")
    public ResponseEntity<List<BatchResponseDTO>> getBatchesByCourse(@PathVariable Integer courseId) {
        try {
            List<Batch> batches = batchService.getAllBatches().stream()
                    .filter(batch -> batch.getCourse() != null &&
                            batch.getCourse().getCourseId().equals(courseId))
                    .filter(Batch::getBatchIsActive)
                    .collect(Collectors.toList());

            List<BatchResponseDTO> responseDTOs = batches.stream()
                    .map(this::convertToResponseDTO)
                    .collect(Collectors.toList());

            return new ResponseEntity<>(responseDTOs, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // ===========================
    // HELPER METHODS - DTO CONVERSION
    // ===========================

    private BatchResponseDTO convertToResponseDTO(Batch batch) {
        BatchResponseDTO dto = new BatchResponseDTO();
        dto.setBatchId(batch.getBatchId());
        dto.setBatchName(batch.getBatchName());
        dto.setBatchStartDate(batch.getBatchStartDate());
        dto.setBatchEndDate(batch.getBatchEndDate());
        dto.setCourseId(batch.getCourse() != null ? batch.getCourse().getCourseId() : null);
        dto.setCourseName(batch.getCourse() != null ? batch.getCourse().getCourseName() : null);
        dto.setPresentationDate(batch.getPresentationDate());
        dto.setCourseFees(batch.getCourseFees());
        dto.setCourseFeesFrom(batch.getCourseFeesFrom());
        dto.setCourseFeesTo(batch.getCourseFeesTo());
        dto.setBatchIsActive(batch.getBatchIsActive());
        dto.setFinalPresentationDate(batch.getFinalPresentationDate());
        dto.setBatchLogoUrl(batch.getBatchLogoUrl());
        dto.setCreatedAt(batch.getCreatedAt());
        dto.setUpdatedAt(batch.getUpdatedAt());
        return dto;
    }

    private Batch convertCreateRequestToEntity(BatchCreateRequestDTO dto) {
        Batch batch = new Batch();
        batch.setBatchName(dto.getBatchName());
        batch.setBatchStartDate(dto.getBatchStartDate());
        batch.setBatchEndDate(dto.getBatchEndDate());

        if (dto.getCourseId() != null) {
            Course course = courseService.getCourseEntityById(dto.getCourseId());
            batch.setCourse(course);
        }

        batch.setPresentationDate(dto.getPresentationDate());
        batch.setCourseFees(dto.getCourseFees());
        batch.setCourseFeesFrom(dto.getCourseFeesFrom());
        batch.setCourseFeesTo(dto.getCourseFeesTo());
        batch.setBatchIsActive(dto.getBatchIsActive());
        batch.setFinalPresentationDate(dto.getFinalPresentationDate());
        batch.setBatchLogoUrl(dto.getBatchLogoUrl());
        return batch;
    }

    private Batch convertUpdateRequestToEntity(BatchUpdateRequestDTO dto) {
        Batch batch = new Batch();
        batch.setBatchName(dto.getBatchName());
        batch.setBatchStartDate(dto.getBatchStartDate());
        batch.setBatchEndDate(dto.getBatchEndDate());

        if (dto.getCourseId() != null) {
            Course course = courseService.getCourseEntityById(dto.getCourseId());
            batch.setCourse(course);
        }

        batch.setPresentationDate(dto.getPresentationDate());
        batch.setCourseFees(dto.getCourseFees());
        batch.setCourseFeesFrom(dto.getCourseFeesFrom());
        batch.setCourseFeesTo(dto.getCourseFeesTo());
        batch.setBatchIsActive(dto.getBatchIsActive());
        batch.setFinalPresentationDate(dto.getFinalPresentationDate());
        batch.setBatchLogoUrl(dto.getBatchLogoUrl());
        return batch;
    }
}