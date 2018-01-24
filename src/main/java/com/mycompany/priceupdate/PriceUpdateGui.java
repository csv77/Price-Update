package com.mycompany.priceupdate;

import java.io.File;
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
import javafx.scene.layout.GridPane;
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
    private Columns[] columnsForPrices;
    private Columns[] columnsForSchema;
    
    @Override
    public void start(Stage primaryStage) {
        BorderPane borderPane = new BorderPane();
        borderPane.setPadding(new Insets(10));
        Button btOpenFile = new Button("Open the excel file");
        
        VBox vBox = new VBox(5);
        vBox.setAlignment(Pos.CENTER);
        
        TextField tfColumnsForPrices = new TextField("J, K, L, M, N, O, P, Q, R, S, T, U");
        Label lbPrices = new Label("Columns of prices ");
        TextField tfColumnsForSchema = new TextField("Z, AA, AB, AC, AD, AE, AF");
        Label lbSchema = new Label("Columns of schema ");
        
        GridPane gridPane = new GridPane();
        gridPane.add(lbPrices, 0, 0);
        gridPane.add(lbSchema, 0, 1);
        gridPane.add(tfColumnsForPrices, 1, 0);
        gridPane.add(tfColumnsForSchema, 1, 1);
        gridPane.setAlignment(Pos.CENTER);
        
        vBox.getChildren().addAll(btOpenFile, gridPane);
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
                    List<String> colsForSchema = textFieldToStringList(tfColumnsForSchema);
                    if(colsForPrices != null && colsForSchema != null) {
                        columnsForPrices = getColumnsFromStringList(colsForPrices);
                        columnsForSchema = getColumnsFromStringList(colsForSchema);
                        Controller controller = new Controller(initialDirectory + file.getName(),
                                initialDirectory + destinationFilename, columnsForPrices, columnsForSchema);
                        lbStatus.setText("File is completed.");
                    }
                    else {
                        lbStatus.setText("Invalid columns.");
                    }
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                }
            }
        });
        
        Scene scene = new Scene(borderPane, 300, 150);
        primaryStage.setTitle("PriceUpload");
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
