package com.mycompany.priceupdate;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SourceWorkbook {
    private Workbook wb;
    private Sheet sheet;
    private List<List<Cell>> listOfPrices = new ArrayList<>();

    public SourceWorkbook(String filename) throws InvalidFormatException, IOException {
        wb = WorkbookFactory.create(new File(filename));
        sheet = wb.getSheetAt(0);
    }

    public List<List<Cell>> getlistOfPrices() {
        return listOfPrices;
    }
    
    public void fillUpListOfPrices() {
        int lastRow = sheet.getLastRowNum();
        for(int rowNum = 1; rowNum <= lastRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            
            int lastColumn = row.getLastCellNum();
            List<Cell> listOfCells = new ArrayList<>();
            for(int column = 0; column < lastColumn; column++) {
                switch(column) {
                    case 0: case 1: case 2: case 3: case 4: case 5: case 6:
                        if(row.getCell(column).toString() != "") {
                            listOfCells.add(row.getCell(column));
                            if(column == 1) {
                                Cell cellOfCurrency = row.createCell(50);
                                cellOfCurrency.setCellValue("FT");
                                listOfCells.add(cellOfCurrency);
                            }
                            else if (column > 1){
                                Cell cellOfCurrency = row.createCell(51);
                                cellOfCurrency.setCellValue("EUR");
                                listOfCells.add(cellOfCurrency);
                            }
                        }
                        break;
                }
            }
            listOfPrices.add(listOfCells);
        }
    }
    
    public static void printListOfList(List<List<Cell>> listOfPrices) {
        for(List<Cell> listOfCells : listOfPrices) {
            for(Cell cell : listOfCells) {
                System.out.print(cell + " ");
            }
            System.out.println();
        }
    }
}
