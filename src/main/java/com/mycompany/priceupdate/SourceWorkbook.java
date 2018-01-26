package com.mycompany.priceupdate;

import java.io.FileInputStream;
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
    private Sheet sheet1;
    private List<String> headerList;
    private List<List<Cell>> listOfPrices = new ArrayList<>();
    private List<List<Cell>> listOfSchema = new ArrayList<>();
    private String filename;
    
    public SourceWorkbook(String filename) throws IOException, InvalidFormatException {
        this.filename = filename;
        headerList = getHeader(filename);
    }

    public String getFilename() {
        return filename;
    }
    
    public List<List<Cell>> getlistOfPrices() {
        return listOfPrices;
    }

    public List<List<Cell>> getListOfSchema() {
        return listOfSchema;
    }

    public List<String> getHeaderList() {
        return headerList;
    }
    
    public List<String> getHeader(String filename) throws IOException, InvalidFormatException {
        List<String> headerList = new ArrayList<>();
        Workbook wb = WorkbookFactory.create(new FileInputStream(filename));
        Sheet sheet = wb.getSheetAt(0);
        boolean notFound = true;
        int i = 0;
        int lastRow = sheet.getLastRowNum();
        while(notFound && i <= lastRow) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0);
            if(cell != null && cell.getStringCellValue().equals("Cikkszám")) {
                notFound = false;
                for(Cell cellInRow : row) {
                    headerList.add(cellInRow.getStringCellValue());
                }
            }
            i++;
        }
        return headerList;
    }
    
    public void fillUpListOfPricesAndListOfSchema() throws IOException, InvalidFormatException {
        wb = WorkbookFactory.create(new FileInputStream(filename));
        sheet1 = wb.getSheetAt(0);
        listOfPrices.clear();
        listOfSchema.clear();
        int lastRow = sheet1.getLastRowNum();
        for(int rowNum = 0; rowNum <= lastRow; rowNum++) {
            Row row = sheet1.getRow(rowNum);
            Cell cell = row.getCell(0);
            if(cell == null || cell.getStringCellValue().equals("Cikkszám") || cell.getStringCellValue().equals("")) {
            	continue;
            }
            
            int lastColumn = row.getLastCellNum();
            int columnForCurrency = lastColumn + 1;
            List<Cell> listOfCellsOfPrices = new ArrayList<>();
            listOfCellsOfPrices.add(cell);
            
            int columnForSchema = lastColumn + 30;
            List<Cell> listOfCellsOfSchema = new ArrayList<>();
            listOfCellsOfSchema.add(cell);
            
            for(Headers header : Headers.values()) {
                String category = header.getCat();
                int index = headerList.indexOf(category);
                if(index < lastColumn) {
                    if(row.getCell(index).getCellTypeEnum().equals(CellType.NUMERIC) && row.getCell(index).getNumericCellValue() == 0) {
                        continue;
                    }
                    if(row.getCell(index).toString().equals("")) {
                        continue;
                    }
                    
                    switch(category) {
                        case "Listaár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.EUR_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Listaár (Ft)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.HU_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Agram konfek. lista ár (HRK)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Agram konfek. transf. ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case "Agram lista ár (HRK)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Agram transfer ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_TR);
                            columnForCurrency += 2;
                            break;
                        case "Okovi konfek. lista ár (SRD)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Okovi konfek. transf. ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case "Okovi lista ár (SRD)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Okovi transfer ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_TR);
                            columnForCurrency += 2;
                            break;
                        case "Konfekcionált ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.EUR_KONF);
                            columnForCurrency += 2;
                            break;
                        case "Konfekcionált ár (Ft)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.HU_KONF);
                            columnForCurrency += 2;
                            break;
                        case "K00001;0;Fuvar %":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.FUVAR);
                            columnForSchema += 2;
                            break;
                        case "K00002;0;Vám %":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.VAM);
                            columnForSchema += 2;
                            break;
                        case "K00003;0;Engedmény %":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.ENGEDMENY);
                            columnForSchema += 2;
                            break;
                        case "K00004;0;Egyéb %":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.EGYEB);
                            columnForSchema += 2;
                            break;
                        case "K00005;0;Hulladék / selejt %":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.HULLADEK);
                            columnForSchema += 2;
                            break;
                        case "K00006;0;Szélesség %":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.SZELESSEG);
                            columnForSchema += 2;
                            break;
                        case "K00007;1;Fix költség":
                            setCellsOfSchemaAndSchemaCode(listOfCellsOfSchema, row, columnForSchema, index, SchemaCat.FIXKTG);
                            columnForSchema += 2;
                            break;
                    }
                    
                }
            }
            listOfPrices.add(listOfCellsOfPrices);
            listOfSchema.add(listOfCellsOfSchema);
        }
    }
    
    private static void setCellsOfCurrencyAndPriceCode(List<Cell> listOfCells, Row row, int lastColumn, int index, PriceCat priceCat){
        listOfCells.add(row.getCell(index));
        Cell cellOfCurrency = row.createCell(lastColumn);
        cellOfCurrency.setCellValue(priceCat.getCurrency());
        listOfCells.add(cellOfCurrency);
        Cell cellOfPriceCode = row.createCell(lastColumn + 1);
        cellOfPriceCode.setCellValue(priceCat.getCode());
        listOfCells.add(cellOfPriceCode);
    }
    
    private static void setCellsOfSchemaAndSchemaCode(List<Cell> listOfCells, Row row, int lastColumn, int index, SchemaCat schemaCat) {
        listOfCells.add(row.getCell(index));
        Cell cellOfSchema = row.createCell(lastColumn);
        cellOfSchema.setCellValue(schemaCat.getType());
        listOfCells.add(cellOfSchema);
        Cell cellOfSchemaCode = row.createCell(lastColumn + 1);
        cellOfSchemaCode.setCellValue(schemaCat.getCode());
        listOfCells.add(cellOfSchemaCode);
    }
    
    public Workbook modifySourceWoorkbook() throws IOException, InvalidFormatException {
        wb = WorkbookFactory.create(new FileInputStream(filename));
        sheet1 = wb.getSheetAt(0);
        int lastRow = sheet1.getLastRowNum();
        int lastCell = headerList.indexOf("Szélesség");
        for(int rowNum = 0; rowNum <= lastRow; rowNum++) {
            Row row = sheet1.getRow(rowNum);
            Cell cell = row.getCell(0);
            if(cell == null || cell.getStringCellValue().equals("Cikkszám") || cell.getStringCellValue().equals("")) {
            	if(cell.getStringCellValue().equals("Cikkszám")) {
                    Cell cellEur = row.createCell(lastCell + 1);
                    cellEur.setCellValue("Beszár EUR");
                }
                continue;
            }
            
            if(cell.getStringCellValue().charAt(0) == '3') {
                Cell cellOfWidth = row.getCell(lastCell);
                if(cellOfWidth != null) {
                    double width = cellOfWidth.getNumericCellValue();
                    if(width != 0) {
                        Cell cellOfSchemaWidth = row.getCell(headerList.indexOf(Headers.SZELESSEG.getCat()));
                        cellOfSchemaWidth.setCellValue(width / 10 - 100);
                    }
                }
            }
            
            Cell cellEur = row.createCell(lastCell + 1);
            int katar = headerList.indexOf("Szállítói utolsó szerződéses ár");
            int fuvar = headerList.indexOf(Headers.FUVAR.getCat());
            int vam = headerList.indexOf(Headers.VAM.getCat());
            int engedmeny = headerList.indexOf(Headers.ENGEDMENY.getCat());
            int egyeb = headerList.indexOf(Headers.EGYEB.getCat());
            int hulladek = headerList.indexOf(Headers.HULLADEK.getCat());
            int szelesseg = headerList.indexOf(Headers.SZELESSEG.getCat());
            int fixktg = headerList.indexOf(Headers.FIXKTG.getCat());
            
            Cell cellKatar = row.getCell(katar);
            Cell cellFuvar = row.getCell(fuvar);
            Cell cellVam = row.getCell(vam);
            Cell cellEngedmeny = row.getCell(engedmeny);
            Cell cellEgyeb = row.getCell(egyeb);
            Cell cellHulladek = row.getCell(hulladek);
            Cell cellSzelesseg = row.getCell(szelesseg);
            Cell cellFixktg = row.getCell(fixktg);
            
            double katarValue = cellKatar.getNumericCellValue();
            double fuvarValue = cellFuvar.getNumericCellValue();
            double vamValue = cellVam.getNumericCellValue();
            double engedmenyValue = cellEngedmeny.getNumericCellValue();
            double egyebValue = cellEgyeb.getNumericCellValue();
            double hulladekValue = cellHulladek.getNumericCellValue();
            double szelessegValue = cellSzelesseg.getNumericCellValue();
            double fixktgValue = cellFixktg.getNumericCellValue();
            
            cellEur.setCellValue(katarValue * (1 + fuvarValue / 100) * (1 + vamValue / 100) *
                    (1 + engedmenyValue / 100) * (1 + egyebValue / 100) * (1 + hulladekValue / 100) *
                    (1 + szelessegValue / 100) + fixktgValue);
        }
        sheet1.autoSizeColumn(lastCell + 1);
        return wb;
    }
}
