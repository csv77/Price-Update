package com.mycompany.priceupdate;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class PriceUpdateGui extends Application {
    private Label lbStatus = new Label();
    private FileChooser chooser = new FileChooser();
    private File file;
    private String initialDirectory;
    private String destinationFilename = "árfelvitel_táblázat.xlsx";
    private Columns[] columns;
    
    @Override
    public void start(Stage primaryStage) {
        BorderPane borderPane = new BorderPane();
        borderPane.setPadding(new Insets(10));
        Button btOpenFile = new Button("Open the excel file");
        
        VBox vBox = new VBox(5);
        vBox.setAlignment(Pos.CENTER);
        
        TextField tfColumnsForPrices = new TextField("J, K, L, M, N, O, P, Q, R, S, T, U");
        vBox.getChildren().addAll(btOpenFile, tfColumnsForPrices);
        borderPane.setTop(vBox);
        
        btOpenFile.setOnAction(e -> {
            lbStatus.setText("");
            chooser.setTitle("Open");
            initialDirectory = LastDirectory.loadThePath();
            chooser.setInitialDirectory(new File(initialDirectory));
            file = chooser.showOpenDialog(primaryStage);
            if(file != null) {
                initialDirectory = file.getParent() + "\\";
                LastDirectory.saveThePath(initialDirectory);
            }
        });
        
        Button btStart = new Button("Start");
        borderPane.setCenter(btStart);
        borderPane.setBottom(lbStatus);
        BorderPane.setAlignment(lbStatus, Pos.CENTER);
        
        btStart.setOnAction(e -> {
            if(file != null) {
                try {
                    List<String> colsForPrices = textFieldToStringList(tfColumnsForPrices);
                    if(colsForPrices != null) {
                        columns = getColumnsFromStringList(colsForPrices);
                        Controller controller = new Controller(initialDirectory + file.getName(),
                                initialDirectory + destinationFilename, columns);
                        lbStatus.setText("File is completed.");
                    }
                    else {
                        lbStatus.setText("Invalid columns.");
                    }
                } catch (FileNotFoundException ex) {
                    lbStatus.setText("File not found.");
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                }
            }
        });
        
        Scene scene = new Scene(borderPane, 250, 150);
        primaryStage.setTitle("Árfelvitelhez");
        primaryStage.setScene(scene);
        primaryStage.show();
    }
    
    private List<String> textFieldToStringList(TextField tfCol) {
        String txt = tfCol.getText();
        String[] columns = txt.toUpperCase().split("[,\\s]");
        List<String> cols = new ArrayList();
        cols.add("A");
        for(String column : columns) {
            if(!column.equals("")) {
                cols.add(column);
            }
        }
        for(String column : cols) {
            for(int i = 0; i < column.length(); i++) {
                if(column.charAt(i) < 'A' || column.charAt(i) > 'Z') {
                    return null;
                }
            }
        }
        return cols;
    }
    
    private Columns[] getColumnsFromStringList(List<String> cols) {
        Columns[] columns = new Columns[cols.size()];
        int i = 0;
        for(String col : cols) {
            columns[i++] = Columns.valueOf(col);
        }
        return columns;
    }

    public static void main(String[] args) {
        launch(args);
    }
}
