package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Receipt;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.Optional;

@Repository
public interface ReceiptRepository extends JpaRepository<Receipt, Integer> {

    Optional<Receipt> findByPaymentId(Integer paymentId);
}