package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Entity
@Table(name = "payment_type_master")
public class PaymentType {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "payment_type_id")
    private Integer paymentTypeId;

    @Column(name = "payment_type_desc", nullable = false, length = 100)
    private String paymentTypeDesc; // "Cheque", "DD", "Bank Transfer", "Cash"
}