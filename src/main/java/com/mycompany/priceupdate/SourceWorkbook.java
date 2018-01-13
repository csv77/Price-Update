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
            
            List<Cell> listOfCells = new ArrayList<>();
            for(Columns column : Columns.values()) {
                switch(column) {
                    case A: case B: case D: case E: case F: case G:
                        if(row.getCell(column.ordinal()) != null) {
                            listOfCells.add(row.getCell(column.ordinal()));
                        }
                        break;
                }
            }
            listOfPrices.add(listOfCells);
        }
    }
}
