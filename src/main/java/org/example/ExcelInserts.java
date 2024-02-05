package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class ExcelInserts {
    public static Map<Integer, String> firstAndSecondInsert(String path) throws IOException {
        Map<Integer, String> map = new HashMap<>();
        FileInputStream file = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(file);


        Sheet sheet = workbook.getSheet("Пакинг");
        Sheet sheet2;
        int raw2counter = 17;
        int raw1counter = 10;
        int mapCounter = 0;
        Row row = sheet.getRow(raw2counter);
        Cell cell = row.getCell(0);


        while (cell.getCellType() != CellType.BLANK) {
            raw1counter++;
            raw2counter++;
            mapCounter++;


            row = sheet.getRow(raw1counter);
            cell = row.getCell(2);
            String str1 = cell.toString();
            cell = row.getCell(11);
            String str2 = cell.toString();
            str2 = str2.replace(".0", "");
            cell = row.getCell(13);
            String str3 = cell.toString();
            str3 = str3.replace(".0", "");

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            sheet2 = workbook.getSheet("Инвойс");
            row = sheet2.getRow(raw2counter);
            cell = row.getCell(11);
            formulaEvaluator.evaluate(cell);
            String str4 = cell.getStringCellValue();
            cell = row.getCell(16);
            formulaEvaluator.evaluate(cell);
            Double num = cell.getNumericCellValue();
            String str5 = String.format("%.2f", num);

            StringBuilder builder = new StringBuilder();
            builder
                    .append(str1)
                    .append(", упакованные в ")
                    .append(str3)
                    .append(" коробок, страна происхождения Китай, количество  ")
                    .append(str2)
                    .append(" шт. Цена за ")
                    .append(str4)
                    .append(" ")
                    .append(str5)
                    .append(" долларов США");


            map.put(mapCounter, String.valueOf(builder));
        }

        int mapDel = map.size();
        map.remove(mapDel);
        return map;
    }

    public static Map<Integer, ArrayList<String>> thirdInsert(String path) throws IOException {

        FileInputStream file = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(file);


        Sheet sheet = workbook.getSheet("Пакинг");
        int i = 1;
        int rawCounter = 10;
        Row raw;
        Map<Integer, ArrayList<String>> map = new HashMap<>();


        while (sheet.getRow(rawCounter).getCell(1).getCellType() != CellType.BLANK) {
            rawCounter++;

            raw = sheet.getRow(rawCounter);
            Cell cell11 = raw.getCell(17);
            if (cell11 != null && cell11.getCellType() != CellType.BLANK) {
                ArrayList<String> list = new ArrayList<String>();
                raw = sheet.getRow(rawCounter);
                Cell cell1 = raw.getCell(0);
                Cell cell2 = raw.getCell(2);
                Cell cell3 = raw.getCell(7);
                Cell cell4 = raw.getCell(10);
                Cell cell5 = raw.getCell(17);
                Cell cell6 = raw.getCell(18);
                Cell cell7 = raw.getCell(19);


                list.add(cell1.toString().replace(".0", ""));
                list.add(cell2.toString());
                list.add(cell3.toString());
                list.add(cell4.toString());
                list.add(cell5.toString());
                list.add(Service.dateMapper(cell6));
                list.add(Service.dateMapper(cell7));
                map.put(i, list);

                i++;
            }
            raw = sheet.getRow(rawCounter);
            Cell cell22 = raw.getCell(20);
            if (cell22 != null && cell22.getCellType() != CellType.BLANK) {
                ArrayList<String> list = new ArrayList<String>();
                raw = sheet.getRow(rawCounter);
                Cell cell1 = raw.getCell(0);
                Cell cell2 = raw.getCell(2);
                Cell cell3 = raw.getCell(7);
                Cell cell4 = raw.getCell(10);
                Cell cell5 = raw.getCell(20);
                Cell cell6 = raw.getCell(21);
                Cell cell7 = raw.getCell(22);


                list.add(cell1.toString().replace(".0", ""));
                list.add(cell2.toString());
                list.add(cell3.toString());
                list.add(cell4.toString());
                list.add(cell5.toString());
                list.add(Service.dateMapper(cell6));
                list.add(Service.dateMapper(cell7));
                map.put(i, list);
                i++;
            }
        }
        return map;
    }

    public static Map<Integer, ArrayList<String>> fourthInsert(String path) throws IOException {

        FileInputStream file = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(file);


        Sheet sheet = workbook.getSheet("Инвойс");
        int i = 1;
        int rawCounter = 17;
        Row raw;
        Map<Integer, ArrayList<String>> map = new HashMap<>();


        while (sheet.getRow(rawCounter).getCell(1).getCellType() != CellType.BLANK) {
            rawCounter++;

            raw = sheet.getRow(rawCounter);
            ArrayList<String> list = new ArrayList<String>();
            raw = sheet.getRow(rawCounter);
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Cell cell1 = raw.getCell(2);
            formulaEvaluator.evaluate(cell1);
            String str1 = cell1.getStringCellValue();

            Cell cell2 = raw.getCell(11);
            formulaEvaluator.evaluate(cell2);
            String str2 = String.valueOf(cell2.getStringCellValue());

            Cell cell3 = raw.getCell(15);
            formulaEvaluator.evaluate(cell3);
            Double dub3 = (cell3.getNumericCellValue());
            String str3 = String.format("%.2f", dub3);


            Cell cell4 = raw.getCell(16);
            formulaEvaluator.evaluate(cell4);
            Double dub4 = (cell4.getNumericCellValue());
            String str4 = String.format("%.2f", dub4);

            String str5 = "50%";


            String formula1 = "ROUNDUP(" + cell4.getAddress() + "*50/100,2)";
            raw.createCell(29, CellType.FORMULA).setCellFormula(formula1);

            Cell cell6 = raw.getCell(29);
            formulaEvaluator.evaluate(cell6);
            String str6 = String.valueOf(formulaEvaluator.evaluate(cell6));
            StringBuilder builder1 = new StringBuilder(str6);
            builder1.delete(0, 39);
            str6 = String.valueOf(builder1);
            str6 = str6.replace("]", "");
            Double dub6 = Double.valueOf((str6));
            str6 = String.format("%.2f", dub6);


            Cell cell7 = raw.getCell(17);
            formulaEvaluator.evaluate(cell7);
            Double dub7 = (cell7.getNumericCellValue());
            String str7 = String.format("%.2f", dub7);


            String formula2 = "ROUNDUP(" + cell7.getAddress() + "*50/100,2)";
            raw.createCell(30, CellType.FORMULA).setCellFormula(formula2);

            Cell cell8 = raw.getCell(30);
            formulaEvaluator.evaluate(cell8);
            String str8 = String.valueOf(formulaEvaluator.evaluate(cell8));
            StringBuilder builder2 = new StringBuilder(str8);
            builder2.delete(0, 39);
            str8 = String.valueOf(builder2);
            str8 = str8.replace("]", "");
            Double dub8 = Double.valueOf(str8);
            str8 = String.format("%.2f", dub8);


            String formula3 = cell7 + "-" + cell8;
            raw.createCell(31, CellType.FORMULA).setCellFormula(formula3);

            Cell cell9 = raw.getCell(31);
            formulaEvaluator.evaluate(cell9);
            String str9 = String.valueOf(formulaEvaluator.evaluate(cell9));
            StringBuilder builder3 = new StringBuilder(str9);
            builder3.delete(0, 39);
            str9 = String.valueOf(builder3);
            str9 = str9.replace("]", "");
            Double dub9 = Double.valueOf(str9);
            str9 = String.format("%.2f", dub9);

            list.add(str1);
            list.add(str2);
            list.add(str3);
            list.add(str4);
            list.add(str5);
            list.add(str6);
            list.add(str7);
            list.add(str8);
            list.add(str9);
            map.put(i, list);

            i++;


        }

        int mapDel = map.size();
        map.remove(mapDel);
        return map;
    }
}
