package com.mycompany.priceupdate;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

public class Controller {
    private SourceWorkbook srWb;
    private DestinationWorkbook dWb;
    private Columns[] columns;
    
    public Controller(String inputFilename, String outputFilename, Columns[] columns) throws InvalidFormatException, IOException {
        this.columns = columns;
        srWb = new SourceWorkbook(inputFilename, columns);
        srWb.fillUpListOfPrices();

        dWb = new DestinationWorkbook(srWb.getlistOfPrices());
        dWb.loadData();

        Workbook wb = dWb.getWb();

        FileOutputStream fileOut = new FileOutputStream(outputFilename);
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }
}
