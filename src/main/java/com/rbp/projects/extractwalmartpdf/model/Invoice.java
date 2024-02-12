package com.rbp.projects.extractwalmartpdf.model;

import lombok.Data;

import java.util.Date;
import java.util.List;

@Data
public class Invoice {
    private String orderNumber;
    private Date date;
    private List<Item> itemList;
    private double subTotal;
    private double savings;
    private double bagFee;
    private double tax;
    private double donation;
    private double driverTip;
    private double total;
    private String cardEndingIn;

    private int dateIndex;

}
