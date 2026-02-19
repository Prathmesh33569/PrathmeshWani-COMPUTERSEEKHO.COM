package com.computerseekho.api.entity;

import jakarta.persistence.*;
import lombok.*;

@Entity
@Table(name = "closure_reason_master")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ClosureReason {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "closure_reason_id")
    private Integer closureReasonId;

    @Column(name = "closure_reason_desc", nullable = false, length = 100)
    private String closureReasonDesc;

    @Column(name = "is_active")
    private Boolean isActive = true;
}