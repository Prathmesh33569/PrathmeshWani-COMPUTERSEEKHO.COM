package com.computerseekho.api.entity;

import jakarta.persistence.*;
import jakarta.validation.constraints.Email;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Size;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.UpdateTimestamp;

import java.time.LocalDateTime;

@Entity
@Table(name = "staff_master",
        uniqueConstraints = {
                @UniqueConstraint(columnNames = "staff_username"),
                @UniqueConstraint(columnNames = "staff_email")
        })
@Data
@NoArgsConstructor
@AllArgsConstructor
public class Staff {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "staff_id")
    private Integer staffId;

    @NotBlank
    @Size(max = 100)
    @Column(name = "staff_name", nullable = false)
    private String staffName;

    @Size(max = 255)
    @Column(name = "photo_url")
    private String photoUrl;

    @NotBlank
    @Column(name = "staff_mobile", nullable = false)
    private String staffMobile;

    @NotBlank
    @Size(max = 100)
    @Email
    @Column(name = "staff_email", nullable = false, unique = true)
    private String staffEmail;

    @NotBlank
    @Size(max = 100)
    @Column(name = "staff_username", nullable = false, unique = true)
    private String staffUsername;

    @NotBlank
    @Size(max = 255)
    @Column(name = "staff_password", nullable = false)
    private String staffPassword;

    @NotBlank
    @Size(max = 50)
    @Column(name = "staff_role", nullable = false)
    private String staffRole;  // "teaching" or "non-teaching"

    @Column(name = "is_active", nullable = false)
    private Boolean isActive = true;

    @Column(name = "staff_bio", columnDefinition = "TEXT")
    private String staffBio;

    @Size(max = 100)
    @Column(name = "staff_designation")
    private String staffDesignation;

    @CreationTimestamp
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @UpdateTimestamp
    @Column(name = "updated_at", nullable = false)
    private LocalDateTime updatedAt;

    public Staff(String staffUsername, String staffEmail, String staffPassword) {
        this.staffUsername = staffUsername;
        this.staffEmail = staffEmail;
        this.staffPassword = staffPassword;
    }
}