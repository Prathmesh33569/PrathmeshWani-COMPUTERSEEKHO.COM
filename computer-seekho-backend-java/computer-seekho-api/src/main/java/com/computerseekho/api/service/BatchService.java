package com.computerseekho.api.service;

import com.computerseekho.api.entity.Batch;

import java.util.List;

public interface BatchService {
    Batch createBatch(Batch batch);

    List<Batch> getAllBatches();

    Batch getBatchById(Integer id);

    Batch updateBatch(Integer id, Batch batch);

    void deleteBatch(Integer id);

    List<Batch> searchBatch(String keyword);
}
