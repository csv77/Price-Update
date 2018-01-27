package com.mycompany.priceupdate;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
            if(cell != null && cell.getStringCellValue().equals(Headers.CIKKSZAM.getCat())) {
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
            if(cell == null || cell.getStringCellValue().equals(Headers.CIKKSZAM.getCat()) || cell.getStringCellValue().equals("")) {
            	continue;
            }
            
            int lastColumn = row.getLastCellNum();
            int columnForCurrency = lastColumn + 20;
            List<Cell> listOfCellsOfPrices = new ArrayList<>();
            listOfCellsOfPrices.add(cell);
            
            int columnForSchema = lastColumn + 40;
            List<Cell> listOfCellsOfSchema = new ArrayList<>();
            listOfCellsOfSchema.add(cell);
            
            for(Headers header : Headers.values()) {
                String category = header.getCat();
                int index = headerList.indexOf(category);
                if(index < lastColumn) {
                    if(row.getCell(index) == null) {
                        continue;
                    }
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
        Sheet sheet2 = wb.getSheetAt(1);
        Row row2 = sheet2.createRow(0);
        
        Cell cellEurRate = row2.createCell(0);
        cellEurRate.setCellValue(315);
        String eurRate = sheet2.getSheetName() + "!" + cellEurRate.getAddress().formatAsString();
        
        Cell cellEurRate2 = row2.createCell(3);
        cellEurRate2.setCellValue(300);
        String eurRate2 = sheet2.getSheetName() + "!" + cellEurRate2.getAddress().formatAsString();
        
        Cell cellUsdRate = row2.createCell(1);
        cellUsdRate.setCellValue(260);
        String usdRate = sheet2.getSheetName() + "!" + cellUsdRate.getAddress().formatAsString();
        
        Cell cellEurUsdRate = row2.createCell(2);
        cellEurUsdRate.setCellValue(1.15);
        String eurUsdRate = sheet2.getSheetName() + "!" + cellEurUsdRate.getAddress().formatAsString();
        
        int lastRow = sheet1.getLastRowNum();
        int lastCell = headerList.indexOf(Headers.SZELESSEG2.getCat());
        for(int rowNum = 0; rowNum <= lastRow; rowNum++) {
            Row row = sheet1.getRow(rowNum);
            Cell cell = row.getCell(0);
            if(cell == null || cell.getStringCellValue().equals(Headers.CIKKSZAM.getCat()) || cell.getStringCellValue().equals("")) {
            	if(cell.getStringCellValue().equals(Headers.CIKKSZAM.getCat())) {
                    Cell cellEur = row.createCell(lastCell + 1);
                    cellEur.setCellValue("Beszár EUR");
                    Cell cellHuf = row.createCell(lastCell + 2);
                    cellHuf.setCellValue("Beszár HUF");
                    Cell cellMarginEur = row.createCell(lastCell + 3);
                    cellMarginEur.setCellValue("Árrés EUR");
                    Cell cellMarginHuf = row.createCell(lastCell + 4);
                    cellMarginHuf.setCellValue("Árrés HUF");
                    Cell cellMarginAgram = row.createCell(lastCell + 5);
                    cellMarginAgram.setCellValue("Árrés Agram");
                    Cell cellMarginOkovi = row.createCell(lastCell + 6);
                    cellMarginOkovi.setCellValue("Árrés Okovi");
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
            Cell cellFt = row.createCell(lastCell + 2);
            Cell cellAgramBeszar = row.getCell(headerList.indexOf(Headers.AGRAM_TR.getCat()));
            Cell cellAgramListaar = row.getCell(headerList.indexOf(Headers.AGRAM_LISTA.getCat()));
            Cell cellOkoviBeszar = row.getCell(headerList.indexOf(Headers.OKOVI_TR.getCat()));
            Cell cellOkoviListaar = row.getCell(headerList.indexOf(Headers.OKOVI_LISTA.getCat()));
            Cell cellEurLista = row.getCell(headerList.indexOf(Headers.EUR_LISTA.getCat()));
            Cell cellHuLista = row.getCell(headerList.indexOf(Headers.HU_LISTA.getCat()));
            
            Cell cellKatar = row.getCell(headerList.indexOf("Szállítói utolsó szerződéses ár"));
            Cell cellDeviza = row.getCell(headerList.indexOf("Szállítói utolsó szerződéses ár pénznem"));
            Cell cellFuvar = row.getCell(headerList.indexOf(Headers.FUVAR.getCat()));
            Cell cellVam = row.getCell(headerList.indexOf(Headers.VAM.getCat()));
            Cell cellEngedmeny = row.getCell(headerList.indexOf(Headers.ENGEDMENY.getCat()));
            Cell cellEgyeb = row.getCell(headerList.indexOf(Headers.EGYEB.getCat()));
            Cell cellHulladek = row.getCell(headerList.indexOf(Headers.HULLADEK.getCat()));
            Cell cellSzelesseg = row.getCell(headerList.indexOf(Headers.SZELESSEG.getCat()));
            Cell cellFixktg = row.getCell(headerList.indexOf(Headers.FIXKTG.getCat()));
            
            String katarPlace = cellKatar.getAddress().formatAsString();
            String fuvarPlace = cellFuvar.getAddress().formatAsString();
            String vamPlace = cellVam.getAddress().formatAsString();
            String engedmenyPlace = cellEngedmeny.getAddress().formatAsString();
            String egyebPlace = cellEgyeb.getAddress().formatAsString();
            String hulladekPlace = cellHulladek.getAddress().formatAsString();
            String szelessegPlace = cellSzelesseg.getAddress().formatAsString();
            String fixktgPlace = cellFixktg.getAddress().formatAsString();
            
            String devizaSzallito = cellDeviza.toString();
            
            double katarValue = cellKatar.getNumericCellValue();
            double fuvarValue = cellFuvar.getNumericCellValue();
            double vamValue = cellVam.getNumericCellValue();
            double engedmenyValue = cellEngedmeny.getNumericCellValue();
            double egyebValue = cellEgyeb.getNumericCellValue();
            double hulladekValue = cellHulladek.getNumericCellValue();
            double szelessegValue = cellSzelesseg.getNumericCellValue();
            double fixktgValue = cellFixktg.getNumericCellValue();
            
            if(devizaSzallito.equals("EUR")) {
                cellEur.setCellFormula("ROUND(" + katarPlace + "*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1+" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)+" + fixktgPlace + ",4)");
                cellFt.setCellFormula("ROUND((" + katarPlace + "*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1+" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)+" + fixktgPlace + ")*" + eurRate + ",4)");
            }
            else if(devizaSzallito.equals("USD")) {
                cellEur.setCellFormula("ROUND(" + katarPlace + "*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1+" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)/" + eurUsdRate + "+" + fixktgPlace + ",4)");
                cellFt.setCellFormula("ROUND(" + katarPlace + "*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1+" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)*" + usdRate + "+" + fixktgPlace + "*" + eurRate + ",4)");
            }
            else if(devizaSzallito.equals("Ft")) {
                cellEur.setCellFormula("ROUND(" + katarPlace + "*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1+" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)/" + eurRate2 + "+" + fixktgPlace + ",4)");
                cellFt.setCellFormula("ROUND(" + katarPlace + "*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1+" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)+" + fixktgPlace + "/" + eurRate + ",4)");
            }
            
            cellAgramBeszar.setCellFormula("ROUND(" + cellEur.getAddress().formatAsString() + "*1.15,4)");
            cellOkoviBeszar.setCellFormula("ROUND(" + cellEur.getAddress().formatAsString() + "*1.1,4)");
            
            cellEur.setCellType(CellType.FORMULA);
            cellFt.setCellType(CellType.FORMULA);
            cellAgramBeszar.setCellType(CellType.FORMULA);
            cellOkoviBeszar.setCellType(CellType.FORMULA);
            cellAgramListaar.setCellType(CellType.FORMULA);
            cellOkoviListaar.setCellType(CellType.FORMULA);
            cellEurLista.setCellType(CellType.FORMULA);
            cellHuLista.setCellType(CellType.FORMULA);
            
            CellStyle style = wb.createCellStyle();
            style.setDataFormat(wb.createDataFormat().getFormat("###,###,##0.0000"));
            cellEur.setCellStyle(style);
            cellAgramBeszar.setCellStyle(style);
            cellOkoviBeszar.setCellStyle(style);
            
            CellStyle style2 = wb.createCellStyle();
            style2.setDataFormat(wb.createDataFormat().getFormat("###,###,##0.00"));
            cellFt.setCellStyle(style2);
        }
        Row row = sheet1.getRow(4);
        if(row != null) {
            int lastCellNum = row.getLastCellNum();
            for(int column = 0; column < lastCellNum; column++) {
                sheet1.autoSizeColumn(column);
            }
        }
        return wb;
    }
}
