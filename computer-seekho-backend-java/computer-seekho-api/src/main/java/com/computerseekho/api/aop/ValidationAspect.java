package com.computerseekho.api.aop;

import com.computerseekho.api.dto.request.EnquiryCreateRequestDTO;
import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

@Aspect
@Component
public class ValidationAspect {

    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    // AOP on createEnquiry method
    @Before("execution(* com.computerseekho.api.service.impl.EnquiryServiceImpl.createEnquiry(..)) && args(dto)")
    public void validateEnquiry(JoinPoint joinPoint, EnquiryCreateRequestDTO dto) {

        logger.info("VALIDATING ENQUIRY DATA...");

        if (dto == null) {
            throw new IllegalArgumentException("Enquiry request cannot be null");
        }

        // Validate name
        if (dto.getEnquirerName() == null || dto.getEnquirerName().trim().isEmpty()) {
            throw new IllegalArgumentException("Enquirer name is required");
        }

        // Validate mobile number (Long)
        Long mobile = dto.getEnquirerMobile();
        if (mobile == null) {
            throw new IllegalArgumentException("Mobile number is required");
        }

        String mobileStr = String.valueOf(mobile);
        if (!mobileStr.matches("\\d{10}")) {
            throw new IllegalArgumentException("Mobile number must be 10 digits");
        }

        // Validate course
        if (dto.getCourseId() == null) {
            throw new IllegalArgumentException("Course is required");
        }

        // Validate staff
        if (dto.getStaffId() == null) {
            throw new IllegalArgumentException("Staff is required");
        }

        logger.info("VALIDATION PASSED");
    }
}
