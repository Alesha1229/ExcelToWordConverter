package org.example;


import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

import static javafx.application.Application.launch;
import static org.example.DocxGenerators.*;
import static org.example.ExcelInserts.*;
import static org.example.Service.dateInsert;
import static org.example.Service.invoisInsert;

public class Main extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("File Chooser Example");

        FileChooser fileChooser = new FileChooser();
        DirectoryChooser directoryChooser = new DirectoryChooser();

        Label inputLabel = new Label("Выберите файл для обработки:");
        TextField inputFileField = new TextField();
        Button chooseInputButton = new Button("Обзор");
        chooseInputButton.setOnAction(e -> {
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                inputFileField.setText(selectedFile.getAbsolutePath());
            }
        });

        Label outputLabel = new Label("Выберите место для сохранения:");
        TextField outputFileField = new TextField();
        Button chooseOutputButton = new Button("Обзор");
        chooseOutputButton.setOnAction(e -> {
            File selectedDirectory = directoryChooser.showDialog(primaryStage);
            if (selectedDirectory != null) {
                outputFileField.setText(selectedDirectory.getAbsolutePath());
            }
        });

        Button processButton = new Button("Обработать файл");
        processButton.setOnAction(e -> {
            String input = inputFileField.getText();
            String output = outputFileField.getText();
            output = output + "\\";
            System.out.println(output);
            String inv;
            try {
                inv = invoisInsert(input);

            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            StringBuilder builder = new StringBuilder(inv);
            builder.setCharAt(5, ' ');
            inv = String.valueOf(builder);
            System.out.println(inv);
            //output = output.replace("1.txt","");
            String output1 = output + "Письмо ЛКВ в Атлант " + inv + ".docx";
            String output2 = output + "Письмо ЛКВ в адрес hs " + inv + ".docx";
            String output3 = output + "БелГим " + inv + ".docx";
            String output4 = output + "Акт скидки " + inv + ".docx";
            String output5 = output + "Письмо Атлант " + inv + ".docx";

            try {
                docxGenerator1(dateInsert(input), invoisInsert(input), firstAndSecondInsert(input), output1);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            try {
                docxGenerator2(dateInsert(input), invoisInsert(input), firstAndSecondInsert(input), output2);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            try {
                docxGenerator3(invoisInsert(input), thirdInsert(input), output3);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            try {
                docxGenerator4(dateInsert(input), invoisInsert(input), fourthInsert(input), output4);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            try {
                docxGenerator5(dateInsert(input), invoisInsert(input), output5);
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
        });

        VBox vBox = new VBox(10);
        vBox.getChildren().addAll(inputLabel, inputFileField, chooseInputButton,
                outputLabel, outputFileField, chooseOutputButton, processButton);

        Scene scene = new Scene(vBox, 400, 250);
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}



