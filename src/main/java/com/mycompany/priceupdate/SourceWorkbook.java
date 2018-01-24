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
    private Columns[] columnsForPrices;
    private List<List<Cell>> listOfSchema = new ArrayList<>();
    private Columns[] columnsForSchema;
    
    public SourceWorkbook(String filename, Columns[] columnsForPrices,
            Columns[] columnsForSchema) throws InvalidFormatException, IOException {
        wb = WorkbookFactory.create(new File(filename));
        sheet = wb.getSheetAt(0);
        this.columnsForPrices = columnsForPrices;
        this.columnsForSchema = columnsForSchema;
    }

    public List<List<Cell>> getlistOfPrices() {
        return listOfPrices;
    }

    public List<List<Cell>> getListOfSchema() {
        return listOfSchema;
    }
    
    public void fillUpListOfPricesAndListOfSchema() {
        int lastRow = sheet.getLastRowNum();
        for(int rowNum = 5; rowNum <= lastRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
                        
            int lastColumn = row.getLastCellNum();
            int columnForCurrency = lastColumn + 1;
            List<Cell> listOfCellsOfPrices = new ArrayList<>();
            for(Columns column : columnsForPrices) {
                if(column.ordinal() < lastColumn) {
                    if(row.getCell(column.ordinal()).getCellTypeEnum().equals(CellType.NUMERIC) && row.getCell(column.ordinal()).getNumericCellValue() == 0) {
                        continue;
                    }
                    if(row.getCell(column.ordinal()).toString().equals("")) {
                        continue;
                    }
                    listOfCellsOfPrices.add(row.getCell(column.ordinal()));
                    switch(column) {
                        case J:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.EUR_LISTA);
                            columnForCurrency += 2;
                            break;
                        case K:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.HU_LISTA);
                            columnForCurrency += 2;
                            break;
                        case L:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.AGRAM_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case M:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.AGRAM_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case N:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.AGRAM_LISTA);
                            columnForCurrency += 2;
                            break;
                        case O:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.AGRAM_TR);
                            columnForCurrency += 2;
                            break;
                        case P:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.OKOVI_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case Q:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.OKOVI_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case R:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.OKOVI_LISTA);
                            columnForCurrency += 2;
                            break;
                        case S:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.OKOVI_TR);
                            columnForCurrency += 2;
                            break;
                        case T:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.EUR_KONF);
                            columnForCurrency += 2;
                            break;
                        case U:
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, PriceCat.HU_KONF);
                            columnForCurrency += 2;
                            break;
                    }
                }
            }
            listOfPrices.add(listOfCellsOfPrices);
            
            int columnForSchema = lastColumn + 30;
            List<Cell> listOfCellsOfSchema = new ArrayList<>();
            for(Columns column : columnsForSchema) {
                if(column.ordinal() < lastColumn) {
                    if(row.getCell(column.ordinal()).toString().equals("")) {
                        continue;
                    }
                    listOfCellsOfSchema.add(row.getCell(column.ordinal()));
                    switch(column) {
                        case Z:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.FUVAR);
                            columnForSchema += 2;
                            break;
                        case AA:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.VAM);
                            columnForSchema += 2;
                            break;
                        case AB:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.ENGEDMENY);
                            columnForSchema += 2;
                            break;
                        case AC:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.EGYEB);
                            columnForSchema += 2;
                            break;
                        case AD:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.HULLADEK);
                            columnForSchema += 2;
                            break;
                        case AE:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.SZELESSEG);
                            columnForSchema += 2;
                            break;
                        case AF:
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, SchemaCat.FIXKTG);
                            columnForSchema += 2;
                            break;
                    }
                }
            }
            listOfSchema.add(listOfCellsOfSchema);
        }
    }
    
    private static void setCellsOfCurrencyAndPriceCode(List<Cell> listOfCells, Row row, int lastColumn, PriceCat priceCat){
        Cell cellOfCurrency = row.createCell(lastColumn);
        cellOfCurrency.setCellValue(priceCat.getCurrency());
        listOfCells.add(cellOfCurrency);
        Cell cellOfPriceCode = row.createCell(lastColumn + 1);
        cellOfPriceCode.setCellValue(priceCat.getCode());
        listOfCells.add(cellOfPriceCode);
    }
    
    private static void setCellsOfSchemaAndSchemaCode(List<Cell> listOfCells, Row row, int lastColumn, SchemaCat schemaCat) {
        Cell cellOfSchema = row.createCell(lastColumn);
        cellOfSchema.setCellValue(schemaCat.getType());
        listOfCells.add(cellOfSchema);
        Cell cellOfSchemaCode = row.createCell(lastColumn + 1);
        cellOfSchemaCode.setCellValue(schemaCat.getCode());
        listOfCells.add(cellOfSchemaCode);
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
