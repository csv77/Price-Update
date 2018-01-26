package com.mycompany.priceupdate;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

public class Controller {
    private SourceWorkbook srWb;
    private DestinationWorkbook dWb;
    private String outputFilename;

   public Controller(String inputFilename, String outputFilename) throws IOException, InvalidFormatException {
        this.outputFilename = outputFilename;
        srWb = new SourceWorkbook(inputFilename);
    }
    
    public void makeDestinationExcelFile() throws InvalidFormatException, IOException {
        srWb.fillUpListOfPricesAndListOfSchema();
        dWb = new DestinationWorkbook(srWb.getlistOfPrices(), srWb.getListOfSchema());

        Workbook wb = dWb.getWb();

        FileOutputStream fileOut = new FileOutputStream(outputFilename);
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }
    
    public void modifiySourceExcelFile() throws IOException, InvalidFormatException {
        Workbook wb = srWb.modifySourceWoorkbook();
        FileOutputStream fileOut = new FileOutputStream(srWb.getFilename());
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }
}
