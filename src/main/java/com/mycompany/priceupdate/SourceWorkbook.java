package com.mycompany.priceupdate;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SourceWorkbook {
    private Workbook wb;
    private Sheet sheet;
    private List<List<Cell>> listOfPrices = new ArrayList<>();
    private Columns[] columns;
    
    public SourceWorkbook(String filename, Columns[] columns) throws InvalidFormatException, IOException {
        wb = WorkbookFactory.create(new File(filename));
        sheet = wb.getSheetAt(0);
        this.columns = columns;
    }

    public List<List<Cell>> getlistOfPrices() {
        return listOfPrices;
    }
    
    public void fillUpListOfPrices() {
        int lastRow = sheet.getLastRowNum();
        for(int rowNum = 5; rowNum <= lastRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
                        
            int lastColumn = row.getLastCellNum();
            int columnForCurrency = lastColumn + 1;
            List<Cell> listOfCells = new ArrayList<>();
            for(Columns column : columns) {
                if(column.ordinal() < lastColumn) {
                    if(row.getCell(column.ordinal()).getCellTypeEnum().equals(CellType.NUMERIC) && row.getCell(column.ordinal()).getNumericCellValue() == 0) {
                        continue;
                    }
                    if(row.getCell(column.ordinal()).toString().equals("")) {
                        continue;
                    }
                    listOfCells.add(row.getCell(column.ordinal()));
                    switch(column) {
                        case J:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.EUR_LISTA);
                            columnForCurrency += 2;
                            break;
                        case K:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.HU_LISTA);
                            columnForCurrency += 2;
                            break;
                        case L:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case M:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case N:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_LISTA);
                            columnForCurrency += 2;
                            break;
                        case O:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.AGRAM_TR);
                            columnForCurrency += 2;
                            break;
                        case P:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case Q:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case R:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_LISTA);
                            columnForCurrency += 2;
                            break;
                        case S:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.OKOVI_TR);
                            columnForCurrency += 2;
                            break;
                        case T:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.EUR_KONF);
                            columnForCurrency += 2;
                            break;
                        case U:
                            setCellsOfCurrencyAndPriceCode(listOfCells, row, columnForCurrency, PriceCat.HU_KONF);
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
    
    private void printListOfPrices() {
        for(List<Cell> cells : listOfPrices) {
            for(Cell cell : cells) {
                System.out.print(cell.toString() + " ");
            }
            System.out.println();
        }
    }
}
