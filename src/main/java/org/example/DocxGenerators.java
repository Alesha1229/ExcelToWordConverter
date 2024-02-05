package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class DocxGenerators {
    public static void docxGenerator1(String date, String serial, Map<Integer, String> map, String savePath) throws IOException {
        ClassLoader classLoader = Main.class.getClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream("doc1.docx");
        XWPFDocument document = new XWPFDocument(inputStream);

        Service.editTextInTable(document, "2222", date + " № " + serial);

        XWPFTable table = document.createTable();
        for (Map.Entry<Integer, String> entry : map.entrySet()) {
            XWPFTableRow row = table.createRow();
            row.createCell().setText(String.valueOf(entry.getKey()));
            row.createCell().setText(entry.getValue());
        }
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontSize(14);
        run.setText("Директор                                                                                                          Д.В. Счастный");


        try {
            document.write(new FileOutputStream(savePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void docxGenerator2(String date, String serial, Map<Integer, String> map, String savePath) throws IOException {
        ClassLoader classLoader = Main.class.getClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream("doc2.docx");
        XWPFDocument document = new XWPFDocument(inputStream);

        Service.editTextInTable(document, "2222", date + " № " + serial);
        Service.editText(document, "3333", "№ " + serial);

        XWPFTable table = document.createTable();
        for (Map.Entry<Integer, String> entry : map.entrySet()) {
            XWPFTableRow row = table.createRow();
            row.createCell().setText(String.valueOf(entry.getKey()));
            row.createCell().setText(entry.getValue());
        }
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontSize(14);
        run.setText("Директор                                                                                                          Д.В. Счастный");


        try {
            document.write(new FileOutputStream(savePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void docxGenerator3(String invois, Map<Integer, ArrayList<String>> map, String savePath) throws IOException {
        XWPFDocument document = new XWPFDocument();

        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontSize(10);
        run.setText("БелГим инвойс " + invois);


        XWPFTable table = document.createTable();
        XWPFTableRow row = table.createRow();
        row.createCell().setText("№ п/п");
        row.createCell().setText("№ по инв");
        row.createCell().setText("Наименование продукции");
        row.createCell().setText("Наименование изготовителя");
        row.createCell().setText("Страна изготовителя");
        row.createCell().setText("№ сертификата / декларации");
        row.createCell().setText("Срок действия c:");
        row.createCell().setText("Срок действия по:");

        for (Map.Entry<Integer, ArrayList<String>> entry : map.entrySet()) {
            XWPFTableRow row2 = table.createRow();
            row2.createCell().setText(String.valueOf(entry.getKey()));
            row2.createCell().setText(entry.getValue().get(0));
            row2.createCell().setText(entry.getValue().get(1));
            row2.createCell().setText(entry.getValue().get(2));
            row2.createCell().setText(entry.getValue().get(3));
            row2.createCell().setText(entry.getValue().get(4));
            row2.createCell().setText(entry.getValue().get(5));
            row2.createCell().setText(entry.getValue().get(6));
        }


        try {
            document.write(new FileOutputStream(savePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void docxGenerator4(String date, String invois, Map<Integer, ArrayList<String>> map, String savePath) throws IOException {
        ClassLoader classLoader = Main.class.getClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream("doc4.docx");
        XWPFDocument document = new XWPFDocument(inputStream);

        Service.editTextBold(document, "4444", "№ " + invois);
        Service.editTextBold(document, "3333", date);

        List<XWPFTable> tables = document.getTables();

        XWPFTable table = tables.get(0);
        //XWPFTable table = document.createTable();

        //XWPFTableRow row = table.createRow();
        table.getRow(0).getCell(0).setText("№ пп");
        table.getRow(0).createCell().setText("Наименование товара");
        table.getRow(0).createCell().setText("Ед. изме- рения");
        table.getRow(0).createCell().setText("Кол-во в ед. изме- рения");
        table.getRow(0).createCell().setText("Перво- началь- ная цена долл. США в ед. измерения");
        table.getRow(0).createCell().setText("Скидка, %");
        table.getRow(0).createCell().setText("Цена со скидкой долл. США в ед. измерения");
        table.getRow(0).createCell().setText("Первона- чальная стоимость долл.США");
        table.getRow(0).createCell().setText("Стоимость со скидкой, долл.США");
        table.getRow(0).createCell().setText("Сумма скидки, долл.США");

        double get6 = 0;
        double get7 = 0;
        double get8 = 0;
        int i = 1;

        for (Map.Entry<Integer, ArrayList<String>> entry : map.entrySet()) {
            //XWPFTableRow row2 = table.createRow();
            table.createRow();
            table.getRow(i).getCell(0).setText(String.valueOf(entry.getKey()));
            table.getRow(i).getCell(1).setText(entry.getValue().get(0));
            table.getRow(i).getCell(2).setText(entry.getValue().get(1));
            table.getRow(i).getCell(3).setText(entry.getValue().get(2));
            table.getRow(i).getCell(4).setText(entry.getValue().get(3));
            table.getRow(i).getCell(5).setText(entry.getValue().get(4));
            table.getRow(i).getCell(6).setText(entry.getValue().get(5));
            table.getRow(i).getCell(7).setText(entry.getValue().get(6));
            String str6 = String.valueOf(entry.getValue().get(6));
            str6 = str6.replace(",", ".");
            double dub6 = Double.parseDouble(str6);
            get6 = get6 + dub6;

            table.getRow(i).getCell(8).setText(entry.getValue().get(7));
            String str7 = String.valueOf(entry.getValue().get(7));
            str7 = str7.replace(",", ".");
            double dub7 = Double.parseDouble(str7);
            get7 = get7 + dub7;

            table.getRow(i).getCell(9).setText(entry.getValue().get(8));
            String str8 = String.valueOf(entry.getValue().get(8));
            str8 = str8.replace(",", ".");
            double dub8 = Double.parseDouble(str8);
            get8 = get8 + dub8;
            i++;
        }
        String str6 = String.format("%.2f", get6);
        String str7 = String.format("%.2f", get7);
        String str8 = String.format("%.2f", get8);

        table.createRow();
        table.getRow(i).getCell(0).setText("");
        table.getRow(i).getCell(1).setText("");
        table.getRow(i).getCell(2).setText("");
        table.getRow(i).getCell(3).setText("");
        table.getRow(i).getCell(4).setText("");
        table.getRow(i).getCell(5).setText("");
        table.getRow(i).getCell(6).setText("Итог:");
        table.getRow(i).getCell(7).setText(str6);
        table.getRow(i).getCell(8).setText(str7);
        table.getRow(i).getCell(9).setText(str8);


        try {
            document.write(new FileOutputStream(savePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void docxGenerator5(String date, String invois, String savePath) throws IOException {
        ClassLoader classLoader = Main.class.getClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream("doc5.docx");
        XWPFDocument document = new XWPFDocument(inputStream);

        Service.editText(document, "4444", "№ " + invois);
        Service.editText(document, "3333", date);

        try {
            document.write(new FileOutputStream(savePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
