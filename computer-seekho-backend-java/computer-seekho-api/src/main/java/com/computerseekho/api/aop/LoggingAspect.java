package com.computerseekho.api.aop;

import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.util.Arrays;

@Aspect
@Component
public class LoggingAspect {

    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    @Before("execution(* com.computerseekho.api.service.impl.*.*(..))")
    public void logBeforeMethod(JoinPoint joinPoint) {
        String methodName = joinPoint.getSignature().getName();
        String className = joinPoint.getTarget().getClass().getSimpleName();
        Object[] args = joinPoint.getArgs();

        logger.info("========================================");
        logger.info("CALLING SERVICE METHOD: {}.{}", className, methodName);
        logger.info("PARAMETERS: {}", Arrays.toString(args));
        logger.info("========================================");
    }

    // ====================================
    // 2. LOG AFTER SERVICE METHOD EXECUTION
    // ====================================
    @AfterReturning(
            pointcut = "execution(* com.computerseekho.api.service.impl.*.*(..))",
            returning = "result"
    )
    public void logAfterMethod(JoinPoint joinPoint, Object result) {
        String methodName = joinPoint.getSignature().getName();
        String className = joinPoint.getTarget().getClass().getSimpleName();

        logger.info("========================================");
        logger.info("SERVICE METHOD COMPLETED: {}.{}", className, methodName);
        logger.info("RETURN VALUE: {}", result);
        logger.info("========================================");
    }

    // ====================================
    // 3. LOG EXECUTION TIME (PERFORMANCE)
    // ====================================
    @Around("execution(* com.computerseekho.api.service.impl.*.*(..))")
    public Object logExecutionTime(ProceedingJoinPoint joinPoint) throws Throwable {
        String methodName = joinPoint.getSignature().getName();
        String className = joinPoint.getTarget().getClass().getSimpleName();

        long startTime = System.currentTimeMillis();

        Object result = joinPoint.proceed(); // execute real method

        long endTime = System.currentTimeMillis();
        long executionTime = endTime - startTime;

        logger.info("{}.{} executed in {} ms", className, methodName, executionTime);

        if (executionTime > 1000) {
            logger.warn("SLOW METHOD DETECTED: {}.{} took {} ms",
                    className, methodName, executionTime);
        }

        return result;
    }

    // ====================================
    // 4. LOG ONLY ENQUIRY SERVICE METHODS
    // ====================================
    @Before("execution(* com.computerseekho.api.service.impl.EnquiryServiceImpl.*(..))")
    public void logEnquiryMethods(JoinPoint joinPoint) {
        logger.info("ENQUIRY SERVICE METHOD CALLED: {}",
                joinPoint.getSignature().getName());
    }

    // ====================================
    // 5. LOG CONTROLLER (API CALLS)
    // ====================================
    @Before("execution(* com.computerseekho.api.controller.*.*(..))")
    public void logAPICall(JoinPoint joinPoint) {
        String methodName = joinPoint.getSignature().getName();
        String className = joinPoint.getTarget().getClass().getSimpleName();

        logger.info("API CALLED: {}.{}", className, methodName);
    }
}
