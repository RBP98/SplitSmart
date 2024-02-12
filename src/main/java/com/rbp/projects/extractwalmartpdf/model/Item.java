package com.rbp.projects.extractwalmartpdf.model;

import lombok.Data;

@Data
public class Item {
    private String name;
    private double amount;
    private int qty;
    private String status;
}
