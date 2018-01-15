package com.mycompany.priceupdate;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.BorderPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class PriceUpdateGui extends Application {
    private Label lbStatus = new Label();
    private FileChooser chooser = new FileChooser();
    private File file;
    private String initialDirectory;
    private String destinationFilename = "árfelvitel_táblázat.xlsx";
    
    @Override
    public void start(Stage primaryStage) {
        BorderPane borderPane = new BorderPane();
        borderPane.setPadding(new Insets(10));
        Button btOpenFile = new Button("Open the excel file");
        borderPane.setTop(btOpenFile);
        BorderPane.setAlignment(btOpenFile, Pos.CENTER);
        
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
                    Controller controller = new Controller(initialDirectory + file.getName(), initialDirectory + destinationFilename);
                    lbStatus.setText("File is completed.");
                } catch (FileNotFoundException ex) {
                    lbStatus.setText("File not found.");
                } catch (InvalidFormatException ex) {
                    lbStatus.setText("Invalid file format.");
                } catch (IOException ex) {
                    lbStatus.setText("Cannot open the file.");
                }
            }
        });
        
        Scene scene = new Scene(borderPane, 250, 100);
        primaryStage.setTitle("Árfelvitelhez");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
