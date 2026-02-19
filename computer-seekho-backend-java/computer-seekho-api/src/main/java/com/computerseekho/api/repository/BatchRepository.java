package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Batch;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface BatchRepository extends JpaRepository<Batch,Integer> {
    List<Batch> findByBatchNameContainingIgnoreCase(String keyword);

    // Optional: get only active batches
    List<Batch> findByBatchIsActiveTrue();
}
