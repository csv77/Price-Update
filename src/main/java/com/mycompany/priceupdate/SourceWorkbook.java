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
                            switch(column) {
                                case 1:
                                    Cell cellOfCurrency = row.createCell(lastColumn + 10);
                                    cellOfCurrency.setCellValue(PriceCat.HU_LISTA.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    Cell cellOfPriceCode = row.createCell(lastColumn + 11);
                                    cellOfPriceCode.setCellValue(PriceCat.HU_LISTA.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 2:
                                    cellOfCurrency = row.createCell(lastColumn + 12);
                                    cellOfCurrency.setCellValue(PriceCat.AGRAM_TR.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 13);
                                    cellOfPriceCode.setCellValue(PriceCat.AGRAM_TR.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 3:
                                    cellOfCurrency = row.createCell(lastColumn + 14);
                                    cellOfCurrency.setCellValue(PriceCat.AGRAM_LISTA.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 15);
                                    cellOfPriceCode.setCellValue(PriceCat.AGRAM_LISTA.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 4:
                                    cellOfCurrency = row.createCell(lastColumn + 16);
                                    cellOfCurrency.setCellValue(PriceCat.OKOVI_TR.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 17);
                                    cellOfPriceCode.setCellValue(PriceCat.OKOVI_TR.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 5:
                                    cellOfCurrency = row.createCell(lastColumn + 18);
                                    cellOfCurrency.setCellValue(PriceCat.OKOVI_LISTA.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 19);
                                    cellOfPriceCode.setCellValue(PriceCat.OKOVI_LISTA.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 6:
                                    cellOfCurrency = row.createCell(lastColumn + 20);
                                    cellOfCurrency.setCellValue(PriceCat.EUR_LISTA.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 21);
                                    cellOfPriceCode.setCellValue(PriceCat.EUR_LISTA.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 7:
                                    cellOfCurrency = row.createCell(lastColumn + 22);
                                    cellOfCurrency.setCellValue(PriceCat.HU_KONF.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 23);
                                    cellOfPriceCode.setCellValue(PriceCat.HU_KONF.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 8:
                                    cellOfCurrency = row.createCell(lastColumn + 24);
                                    cellOfCurrency.setCellValue(PriceCat.AGRAM_KONF_TR.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 25);
                                    cellOfPriceCode.setCellValue(PriceCat.AGRAM_KONF_TR.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 9:
                                    cellOfCurrency = row.createCell(lastColumn + 26);
                                    cellOfCurrency.setCellValue(PriceCat.AGRAM_KONF_LISTA.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 27);
                                    cellOfPriceCode.setCellValue(PriceCat.AGRAM_KONF_LISTA.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 10:
                                    cellOfCurrency = row.createCell(lastColumn + 28);
                                    cellOfCurrency.setCellValue(PriceCat.OKOVI_KONF_TR.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 29);
                                    cellOfPriceCode.setCellValue(PriceCat.OKOVI_KONF_TR.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 11:
                                    cellOfCurrency = row.createCell(lastColumn + 30);
                                    cellOfCurrency.setCellValue(PriceCat.OKOVI_KONF_LISTA.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 31);
                                    cellOfPriceCode.setCellValue(PriceCat.OKOVI_KONF_LISTA.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
                                case 12:
                                    cellOfCurrency = row.createCell(lastColumn + 32);
                                    cellOfCurrency.setCellValue(PriceCat.EUR_KONF.getCurrency());
                                    listOfCells.add(cellOfCurrency);
                                    cellOfPriceCode = row.createCell(lastColumn + 33);
                                    cellOfPriceCode.setCellValue(PriceCat.EUR_KONF.getCode());
                                    listOfCells.add(cellOfPriceCode);
                                    break;
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
