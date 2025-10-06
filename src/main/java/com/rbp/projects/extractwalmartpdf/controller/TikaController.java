package com.rbp.projects.extractwalmartpdf.controller;

import com.rbp.projects.extractwalmartpdf.model.Invoice;
import com.rbp.projects.extractwalmartpdf.model.Item;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.pdf.PDFParser;
import org.apache.tika.sax.BodyContentHandler;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
@RestController
public class TikaController {
    static Invoice invoice;
    static int numberOfItems;

    static String orderNumber;
    @PostMapping("/upload")
    public Invoice uploadFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            throw new IllegalArgumentException("Please select a file to upload.");
        }

        try {
            // Process the uploaded PDF file
            List<String> lines = readPdf(file);

            // Create the Invoice object
            invoice = createInvoice(lines);

            return invoice;
        } catch (Exception e) {
            throw new RuntimeException("Failed to upload file.", e);
        }
    }

    private List<String> readPdf(MultipartFile file) throws IOException, SAXException, TikaException {
        List<String> lines = new ArrayList<>();
        BodyContentHandler contentHandler = new BodyContentHandler();

        // Read PDF file content
        try (InputStream fileInputStream = file.getInputStream()) {
            Metadata metadata = new Metadata();
            ParseContext context = new ParseContext();
            PDFParser pdfParser = new PDFParser();
            pdfParser.parse(fileInputStream, contentHandler, metadata, context);
            String[] content = contentHandler.toString().split("\\r?\\n");
            for (String line : content) {
                if (!line.isEmpty()) {
                    lines.add(line);
                }
            }
        }
        return lines;
    }

    public static Invoice createInvoice(List<String> strings) {
        Invoice invoice = new Invoice();
        try{
            SimpleDateFormat dateFormat = new SimpleDateFormat("MMM dd, yyyy");
            int index = 0;
            for (String s : strings) {
                if(isDate(s) && invoice.getDate()==null){
                    Date date = dateFormat.parse(s);
                    invoice.setDate(date);
                    invoice.setDateIndex(index);
                }else if(isOrderNumber(s) && invoice.getOrderNumber()== null){
                    invoice.setOrderNumber(s.substring(7));
                    orderNumber = invoice.getOrderNumber();
                }else if(isSubTotal(s) && invoice.getSubTotal()==0.0){
                    invoice.setSubTotal(getAmountFromString(s));
                }else if(isSavings(s) && invoice.getSavings() == 0.0){
                    invoice.setSavings(-getAmountFromString(s));
                }else if(isBagFee(s) && invoice.getBagFee() == 0.0){
                    invoice.setBagFee(getAmountFromString(s));
                }else if(isDonation(s) && invoice.getDonation() == 0.0){
                    invoice.setDonation(getAmountFromString(s));
                }else if(isTotal(s) && invoice.getTotal() == 0.0){
                        invoice.setTotal(getAmountFromString(s));
                }else if(isDriverTip(s) && invoice.getDriverTip() == 0.0){
                    invoice.setDriverTip(getAmountFromString(s));
                } else if(isTax(s) && invoice.getTax() == 0.0) {
                    invoice.setTax(getAmountFromString(s));
                } else if (isCardNumber(s) && invoice.getCardEndingIn() == null) {
                    invoice.setCardEndingIn((s.trim()).substring(s.length() - 4));
                }
                index++;
            }
            List<String> rawItemList = getRawItemListFromStrings(strings, invoice);

            List<Item> itemList = createItemList(rawItemList);

            invoice.setItemList(itemList);
            numberOfItems = itemList.size();

//            int count = 0;
//        for (RowsInInvoice row : rowsList) {
//            System.out.println(count + " " +row.getContent());
//            count++;
//        }
//        for(String s: rawItemList){
//            System.out.println(count + " " +s);
//            count++;
//        }
//
//        System.out.println(invoice.getDateIndex());
//
//        System.out.println(invoice.getOrderNumber());
//
//        System.out.println(invoice.getDate());
//        System.out.println(invoice.getSubTotal());
//        System.out.println(invoice.getTotal());
//        System.out.println(invoice.getTax());
//        System.out.println(invoice.getSavings());
//        System.out.println(invoice.getBagFee());
//        System.out.println(invoice.getDonation());
//        System.out.println(invoice.getCardEndingIn());
//        System.out.println(invoice.getDriverTip());

        }
        catch(ParseException parseException){
            parseException.printStackTrace();
        }

        return invoice;

    }

    private static List<Item> createItemList(List<String> rawItemList) {
        List<Item> finalItemsInInvoice = new ArrayList<>();
        for(String s: rawItemList){
            Item item = new Item();
            item.setName(getItemNameFromString(s));
            item.setStatus(getStatusFromString(s));
            item.setQty(getQuantityFromString(s));
            item.setAmount(getAmountFromString(s));
            finalItemsInInvoice.add(item);
        }

        return finalItemsInInvoice;
    }

    private static List<String> getRawItemListFromStrings(List<String> strings, Invoice invoice) {
        List<String> rawItemList = new ArrayList<>();
        for (int i = 0; i < invoice.getDateIndex(); i++) {
            if (!strings.get(i).contains("$")) {
                if (i != 0) {
                    rawItemList.add((((strings.get(i).concat(" ")).concat(strings.get(i + 1))).concat(" ")).concat(strings.get(i + 2)));
                    i += 2;
                }
            } else {
                rawItemList.add(strings.get(i));
            }
        }
        return rawItemList;
    }

    public static boolean isDate(String s){
        if(s.startsWith("Jan ")
                || s.startsWith("Feb ")
                || s.startsWith("Mar ")
                || s.startsWith("Apr ")
                || s.startsWith("May ")
                || s.startsWith("Jun ")
                || s.startsWith("Jul ")
                || s.startsWith("Aug ")
                || s.startsWith("Sep ")
                || s.startsWith("Oct ")
                || s.startsWith("Nov ")
                || s.startsWith("Dec ")){
            return true;
        }
        return false;
    }

    public static boolean isOrderNumber(String s){
        return s.startsWith("Order#");
    }

    public static boolean isSubTotal(String s){
        return s.startsWith("Subtotal");
    }

    public static boolean isDriverTip(String s){
        return s.startsWith("Driver tip");
    }

    public static boolean isTax(String s){
        return s.startsWith("Tax");
    }

    public static boolean isBagFee(String s){
        return s.startsWith("Bag fee");
    }

    public static boolean isSavings(String s){
        return s.startsWith("Savings");
    }

    public static boolean isCardNumber(String s){
        return s.startsWith("Ending in");
    }
    public static boolean isHyperLink(String s){
        return s.startsWith("https://") || s.startsWith("http://");
    }
    public static boolean isDonation(String s){
        return s.startsWith("Donation ");
    }
    public static boolean isTotal(String s){
        return s.startsWith("Total ");
    }
    public static double getAmountFromString(String s) {
        Pattern pattern = Pattern.compile("\\$(-?\\d+\\.\\d+)");
        Matcher matcher = pattern.matcher(s);
        if (matcher.find()) {
            return Double.parseDouble(matcher.group(1));
        }
        return 0.0;
    }

    public static String getStatusFromString(String str) {
        Pattern pattern = Pattern.compile("(Unavailable|Shopped|Weight-adjusted|You’re all set! No need to return this item|Return complete)");
        Matcher matcher = pattern.matcher(str);

        if (matcher.find()) {
            return matcher.group(1);
        } else {
            return "";
        }
    }

    public static int getQuantityFromString(String s){
        Pattern pattern = Pattern.compile("Qty (\\d+)");
        Matcher matcher = pattern.matcher(s);
        if(matcher.find()){
            return Integer.parseInt(matcher.group(1));
        }
        else{
            return 0;
        }
    }

    public static String getItemNameFromString(String s){
        Pattern pattern = Pattern.compile("^(.*)(Unavailable|Shopped|Weight-adjusted|You’re all set! No need to return this item|Return Complete)");
        Matcher matcher = pattern.matcher(s);
        if(matcher.find()){
            return matcher.group(1).trim();
        }
        else{
            pattern = Pattern.compile("^(.*)(Qty)");
            matcher = pattern.matcher(s);
            if(matcher.find()){
                return matcher.group(1).trim();
            }
            else{
                return "Unable to get Item Name";
            }
        }
    }

}


