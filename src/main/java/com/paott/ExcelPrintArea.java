package com.paott;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPrintArea {

    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // セルにデータを入力 (例)
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("A1");
        row1.createCell(1).setCellValue("B1");
        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("A2");
        row2.createCell(1).setCellValue("B2");
        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue("A3");
        row3.createCell(1).setCellValue("B3");

        // 印刷範囲を設定 (A1:B2)
        workbook.setPrintArea(0, "$A$1:$B$2"); // シートのインデックス、印刷範囲の文字列

        // Excelファイルを出力
        try (FileOutputStream fileOut = new FileOutputStream("print_area.xlsx")) {
            workbook.write(fileOut);
        }

        workbook.close();
    }
}
