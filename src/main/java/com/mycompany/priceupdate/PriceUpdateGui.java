package com.mycompany.priceupdate;

import java.io.File;
import java.io.IOException;
import javafx.application.Application;
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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class PriceUpdateGui extends Application {
    private Label lbStatus = new Label();
    private FileChooser chooser = new FileChooser();
    private File file;
    private String initialDirectory;
    private String destinationFilename = "árfelvitel_táblázat.xlsx";
    private Controller controller;
    private TextField tfEUR;
    private TextField tfUSD;
    
    @Override
    public void start(Stage primaryStage) {
        VBox vBox = new VBox(5);
        vBox.setPadding(new Insets(10));
        vBox.setAlignment(Pos.CENTER);
        
        Button btOpenFile = new Button("Open the excel file");
        btOpenFile.setOnAction(e -> {
            lbStatus.setText("");
            chooser.setTitle("Open");
            initialDirectory = LastDirectory.loadThePath();
            chooser.setInitialDirectory(new File(initialDirectory));
            file = chooser.showOpenDialog(primaryStage);
            if(file != null) {
                initialDirectory = file.getParent() + "\\";
                LastDirectory.saveThePath(initialDirectory);
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
        
        GridPane gridPane = new GridPane();
        tfEUR = new TextField();
        tfUSD = new TextField();
        gridPane.add(new Label("Ft/EUR rate: "), 0, 0);
        gridPane.add(new Label("Ft/USD rate: "), 0, 1);
        gridPane.add(tfEUR, 1, 0);
        gridPane.add(tfUSD, 1, 1);
        gridPane.setAlignment(Pos.CENTER);
        
        Button btCreateDestinationExcelFile = new Button("Create PriceUpload excel");
        Button btCreateModifiedSourceExcelFile = new Button("Calculate width and EUR purchasing price");
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
                    controller.modifiySourceExcelFile();
                    lbStatus.setText("New source excel is finished.");
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                }
            }
        });
        
        Scene scene = new Scene(vBox, 300, 180);
        primaryStage.setTitle("PriceUpload");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
