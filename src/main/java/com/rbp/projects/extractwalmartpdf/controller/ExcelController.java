package com.rbp.projects.extractwalmartpdf.controller;

import com.rbp.projects.extractwalmartpdf.model.Invoice;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.Tika;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

@RestController
public class ExcelController {

    private static final String TEMPLATE_PATH = "src/main/resources/ExcelTemplate.xlsx";
//    private static final String OUTPUT_PATH = "src/main/resources/" + TikaController.orderNumber + ".xlsx";
//    private static final String DOWNLOAD_FILENAME = TikaController.orderNumber + ".xlsx";

    @PostMapping("/downloadInvoice")
    public ResponseEntity<InputStreamResource> downloadInvoiceInSpreadsheet() throws IOException {
        try (InputStream file = new FileInputStream(TEMPLATE_PATH);
             XSSFWorkbook workbook = new XSSFWorkbook(file)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            Cell cell;

            int n = TikaController.numberOfItems;
            Invoice invoice = TikaController.invoice;
            for (int i = 1; i <= n; i++) {
                cell = sheet.getRow(i).getCell(0);
                cell.setCellValue(((invoice.getItemList()).get(i - 1)).getName());
                cell = sheet.getRow(i).getCell(1);
                cell.setCellValue(((invoice.getItemList()).get(i - 1)).getQty());
                cell = sheet.getRow(i).getCell(2);
                cell.setCellValue(((invoice.getItemList()).get(i - 1)).getAmount());
            }

            cell = sheet.getRow(37).getCell(1);
            cell.setCellValue(invoice.getBagFee());

            cell = sheet.getRow(38).getCell(1);
            cell.setCellValue(invoice.getDonation());

            cell = sheet.getRow(39).getCell(1);
            cell.setCellValue(invoice.getDriverTip());

            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

            try (FileOutputStream outputStream = new FileOutputStream("src/main/resources/" + TikaController.orderNumber + ".xlsx")) {
                workbook.write(outputStream);
            }

            InputStreamResource resource = new InputStreamResource(new FileInputStream("src/main/resources/" + TikaController.orderNumber + ".xlsx"));

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + TikaController.orderNumber + ".xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(new File("src/main/resources/" + TikaController.orderNumber + ".xlsx").length())
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(resource);
        } catch (IOException e) {
            return ResponseEntity.status(500).body(null);
        }
    }
}
