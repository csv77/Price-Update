package com.mycompany.priceupdate;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javafx.application.Application;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class PriceUpdateGui extends Application {
    private Label lbStatus = new Label();
    private FileChooser chooser = new FileChooser();
    private File file;
    private String initialDirectory;
    private String destinationFilename = "árfelvitel_táblázat.xlsx";
    private Controller controller;
    private TextField[] tfRates = new TextField[6];

    @Override
    public void start(Stage primaryStage) {
    	VBox vBox = new VBox(5);
        vBox.setPadding(new Insets(10));
        vBox.setAlignment(Pos.CENTER);
        
        GridPane gridPane = new GridPane();
        gridPane.add(new Label("HUF/EUR rate: "), 0, 0);
        gridPane.add(new Label("HUF/EUR rate for divide: "), 0, 1);
        gridPane.add(new Label("HUF/USD rate: "), 0, 2);
        gridPane.add(new Label("USD/EUR rate: "), 2, 0);
        gridPane.add(new Label("KN/EUR rate: "), 2, 1);
        gridPane.add(new Label("DIN/EUR rate: "), 2, 2);
        gridPane.setAlignment(Pos.CENTER);
        gridPane.setVgap(2);
        gridPane.setHgap(2);
        
        for(int i = 0; i < tfRates.length; i++) {
            tfRates[i] = new TextField();
            tfRates[i].setPrefColumnCount(5);
            if(i < 3) {
                gridPane.add(tfRates[i], 1, i % 3);
            }
            else {
                gridPane.add(tfRates[i], 3, i % 3);
            }
            switch(i) {
                case 0:
                    tfRates[i].setText("328");
                    break;
                case 1:
                	tfRates[i].setText("300");
                    break;
                case 2:
                    tfRates[i].setText("285");
                    break;
                case 3:
                    tfRates[i].setText("1.15");
                    break;
                case 4:
                    tfRates[i].setText("7.8");
                    break;
                case 5:
                    tfRates[i].setText("123");
                    break;
            }
            int j = i;
            tfRates[i].textProperty().addListener(new ChangeListener<String>() {
                @Override
                public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                    if(!newValue.matches("\\d{0,4}([\\.]\\d{0,4})?")) {
                        tfRates[j].setText(oldValue);
                    }
                }
            });
        }
        
        Button btOpenFile = new Button("Open the excel file");
        btOpenFile.setOnAction(e -> {
            lbStatus.setText("");
            chooser.setTitle("Open");
            initialDirectory = DataToSave.loadThePath();
            chooser.setInitialDirectory(new File(initialDirectory));
            file = chooser.showOpenDialog(primaryStage);
            if(file != null) {
                initialDirectory = file.getParent() + "\\";
                DataToSave.saveThePath(initialDirectory);
                try {
                    controller = new Controller(initialDirectory + file.getName(),
                            initialDirectory + destinationFilename);
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                }
            }
        });
        
        Button btCreateDestinationExcelFile = new Button("Create PriceUpload excel");
        Button btCreateModifiedSourceExcelFile = new Button("Set formulas and calculate purchasing prices");
        vBox.getChildren().addAll(btOpenFile, gridPane, btCreateModifiedSourceExcelFile, btCreateDestinationExcelFile, lbStatus);
        
        btCreateDestinationExcelFile.setOnAction(e -> {
            if(controller != null) {
                try {
                    controller.makeDestinationExcelFile();
                    lbStatus.setText("PriceUpload excel is finished.");
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                }
            }
        });
        
        btCreateModifiedSourceExcelFile.setOnAction(e -> {
            if(controller != null) {
                try {
                    controller.setRates(getDoublesFromTextFields());
                    controller.modifiySourceExcelFile();
                    lbStatus.setText("New source excel is finished.");
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                }
            }
        });
        
        Scene scene = new Scene(vBox, 400, 200);
        primaryStage.setTitle("PriceUpload");
        primaryStage.setScene(scene);
        primaryStage.show();
    }
    
    private Double[] getDoublesFromTextFields() {
        Double[] rates = new Double[tfRates.length];
        int i = 0;
        for(TextField tfRate : tfRates) {
            String rate = tfRate.getText();
            if(rate.equals("")) {
                rates[i] = new Double(0);
            }
            else {
                rates[i] = Double.parseDouble(rate);
            }
            i++;
        }
        return rates;
    }

    public static void main(String[] args) {
    	launch(args);
    }
}
