package com.mycompany.priceupdate;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

public class Controller {
    private SourceWorkbook srWb;
    private DestinationWorkbook dWb;
    
    public Controller(String inputFilename, String outputFilename, Columns[] columnsForPrices,
            Columns[] columnsForSchema) throws InvalidFormatException, IOException {
        srWb = new SourceWorkbook(inputFilename, columnsForPrices, columnsForSchema);
        srWb.fillUpListOfPricesAndListOfSchema();
        
        dWb = new DestinationWorkbook(srWb.getlistOfPrices(), srWb.getListOfSchema());
        
        Workbook wb = dWb.getWb();

        FileOutputStream fileOut = new FileOutputStream(outputFilename);
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }
}
