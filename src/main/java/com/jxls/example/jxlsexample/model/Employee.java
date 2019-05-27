package com.jxls.example.jxlsexample.model;

import lombok.Builder;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author Leonid Vysochyn
 */
@Builder
@Data
public class Employee {
    private String name;
    private String birthDate;
    private Long payment;
    private Float bonus;
}
