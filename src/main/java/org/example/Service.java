package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;

public class Service {
    public static String dateMapper(Cell cellDate){
        //System.out.println(inputDate);
        //inputDate = inputDate.replace(" ", "");
        //System.out.println(inputDate);
        if (cellDate.getCellType() == CellType.NUMERIC) {
            String inputDate = cellDate.toString();
            DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy", new Locale("ru"));
            LocalDate parsedDate = LocalDate.parse(inputDate, inputFormatter);
            DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
            String outputDate = parsedDate.format(outputFormatter);
            return outputDate;
        } else if (cellDate.getCellType() == CellType.STRING) {
            String inputDate = cellDate.toString();
            inputDate = inputDate.replace(" ", "");
            return inputDate;
        }


        return null;
    }
    static void editTextBold(XWPFDocument document, String oldText, String newText) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null && text.contains(oldText)) {
                    run.setBold(true);
                    text = text.replace(oldText, newText);
                    run.setText(text, 0);
                }
            }
        }
    }

    static void editText(XWPFDocument document, String oldText, String newText) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null && text.contains(oldText)) {
                    text = text.replace(oldText, newText);
                    run.setText(text, 0);
                }
            }
        }
    }

    static void editTextInTable(XWPFDocument document, String oldText, String newText) {
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    editTextInCell(cell, oldText, newText);
                }
            }
        }
    }
    private static void editTextInCell(XWPFTableCell cell, String oldText, String newText) {
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);

                if (text != null && text.contains(oldText)) {
                    text = text.replace(oldText, newText);
                    run.setText(text, 0);
                }
            }
        }
    }
    public static String dateInsert(String path) throws IOException {
        FileInputStream file = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheet("Пакинг");
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(3);
        String date = cell.toString();
        StringBuilder builder = new StringBuilder(date);
        builder.delete(0,6);
        date = String.valueOf(builder);
        return date;
    }
    public static String invoisInsert(String path) throws IOException {
        FileInputStream file = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheet("Пакинг");
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(2);
        String invois = cell.toString();
        return invois;
    }
}
