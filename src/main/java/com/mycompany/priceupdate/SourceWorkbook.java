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
    private int[] columns = {0, 1, 2, 3, 4, 5, 6};
    
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
            int columnForCurrency = lastColumn + 1;
            List<Cell> listOfCells = new ArrayList<>();
            for(int column : columns) {
                if(column < lastColumn && !row.getCell(column).toString().equals("")) {
                    listOfCells.add(row.getCell(column));
                    switch(column) {
                        case 1:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.HU_LISTA);
                            columnForCurrency += 2;
                            break;
                        case 2:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_TR);
                            columnForCurrency += 2;
                            break;
                        case 3:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_LISTA);
                            columnForCurrency += 2;
                            break;
                        case 4:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_TR);
                            columnForCurrency += 2;
                            break;
                        case 5:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_LISTA);
                            columnForCurrency += 2;
                            break;
                        case 6:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.EUR_LISTA);
                            columnForCurrency += 2;
                            break;
                        case 7:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.HU_KONF);
                            columnForCurrency += 2;
                            break;
                        case 8:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case 9:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case 10:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case 11:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case 12:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.EUR_KONF);
                            columnForCurrency += 2;
                            break;
                    }
                }
            }
            listOfPrices.add(listOfCells);
        }
    }
    
    public static void setCellsOfCurrencyAndPriceCode(List<Cell> listOfCells, Row row, int lastColumn, PriceCat priceCat){
        Cell cellOfCurrency = row.createCell(lastColumn);
        cellOfCurrency.setCellValue(priceCat.getCurrency());
        listOfCells.add(cellOfCurrency);
        Cell cellOfPriceCode = row.createCell(lastColumn + 1);
        cellOfPriceCode.setCellValue(priceCat.getCode());
        listOfCells.add(cellOfPriceCode);
    }
}
