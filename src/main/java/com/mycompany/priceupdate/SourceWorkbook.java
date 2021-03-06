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
    private Double[] rates;
    
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

    public void setRates(Double[] rates) {
        this.rates = rates;
    }
    
    private List<String> getHeader(String filename) throws IOException, InvalidFormatException {
        List<String> headerList = new ArrayList<>();
        Workbook wb = WorkbookFactory.create(new FileInputStream(filename));
        Sheet sheet = wb.getSheetAt(0);
        boolean notFound = true;
        int i = 0;
        int lastRow = sheet.getLastRowNum();
        while(notFound && i <= lastRow) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0);
            if(cell != null && cell.toString().equals(Headers.CIKKSZAM.getCat())) {
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
        int lastCell = headerList.size() - 1;
        for(int rowNum = 0; rowNum <= lastRow; rowNum++) {
            Row row = sheet1.getRow(rowNum);
            Cell cell = row.getCell(0);
            if(cell == null || cell.toString().equals(Headers.CIKKSZAM.getCat()) || cell.toString().equals("")) {
            	if(cell != null && cell.toString().equals(Headers.CIKKSZAM.getCat()) && !headerList.contains(AddedHeaders.BESZAREUR.getCat())) {
                    int i = 1;
                    for(AddedHeaders addedHeaders : AddedHeaders.values()) {
                        Cell addedCell = row.createCell(lastCell + i);
                        addedCell.setCellValue(addedHeaders.getCat());
                        headerList.add(addedCell.getStringCellValue());
                        i++;
                    }
                }
                continue;
            }
            
            int lastColumn = row.getLastCellNum();
            int columnForCurrency = lastColumn + 100;
            List<Cell> listOfCellsOfPrices = new ArrayList<>();
            listOfCellsOfPrices.add(cell);
            
            int columnForSchema = lastColumn + 200;
            List<Cell> listOfCellsOfSchema = new ArrayList<>();
            listOfCellsOfSchema.add(cell);
            
            for(String header : headerList) {
                int index = headerList.indexOf(header);
                if(index < lastColumn) {
                    if(row.getCell(index) == null) {
                        continue;
                    }
                    if((row.getCell(index).getCellTypeEnum().equals(CellType.NUMERIC) || row.getCell(index).getCellTypeEnum().equals(CellType.FORMULA))
                    		&& row.getCell(index).getNumericCellValue() == 0) {
                        continue;
                    }
                    if(row.getCell(index).toString().equals("")) {
                        continue;
                    }
                    
                    switch(header) {
                        case "Új listaár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.EUR_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Új listaár (Ft)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.HU_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Új Agram konfek. lista ár (HRK)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Új Agram konfek. transf. ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case "Új Agram lista ár (HRK)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Új Agram transfer ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.AGRAM_TR);
                            columnForCurrency += 2;
                            break;
                        case "Új Okovi konfek. lista ár (SRD)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_KONF_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Új Okovi konfek. transf. ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_KONF_TR);
                            columnForCurrency += 2;
                            break;
                        case "Új Okovi lista ár (SRD)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_LISTA);
                            columnForCurrency += 2;
                            break;
                        case "Új Okovi transfer ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.OKOVI_TR);
                            columnForCurrency += 2;
                            break;
                        case "Új konfekcionált ár (EUR)":
                            setCellsOfCurrencyAndPriceCode(listOfCellsOfPrices, row, columnForCurrency, index, PriceCat.EUR_KONF);
                            columnForCurrency += 2;
                            break;
                        case "Új konfekcionált ár (Ft)":
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
                            Cell cellOfDiscount = row.getCell(index);
                            double oldValue = cellOfDiscount.getNumericCellValue();
                            cellOfDiscount.setCellValue(100 - oldValue);
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
    
    public Workbook addFormulasToSourceWoorkbook() throws IOException, InvalidFormatException {
        wb = WorkbookFactory.create(new FileInputStream(filename));
        sheet1 = wb.getSheetAt(0);
        Sheet sheet2;
        if(wb.getSheet("CurrencyRates") == null) {
        	sheet2 = wb.createSheet("CurrencyRates");
        }
        else {
        	sheet2 = wb.getSheet("CurrencyRates");
        }
        Row row2 = sheet2.createRow(0);
        
        Cell cellEurRate = row2.createCell(0);
        cellEurRate.setCellValue(rates[0]);
        String eurRate = sheet2.getSheetName() + "!" + cellEurRate.getAddress().formatAsString();
        
        Cell cellEurRate2 = row2.createCell(3);
        cellEurRate2.setCellValue(rates[1]);
        String eurRate2 = sheet2.getSheetName() + "!" + cellEurRate2.getAddress().formatAsString();
        
        Cell cellUsdRate = row2.createCell(1);
        cellUsdRate.setCellValue(rates[2]);
        String usdRate = sheet2.getSheetName() + "!" + cellUsdRate.getAddress().formatAsString();
        
        Cell cellEurUsdRate = row2.createCell(2);
        cellEurUsdRate.setCellValue(rates[3]);
        String eurUsdRate = sheet2.getSheetName() + "!" + cellEurUsdRate.getAddress().formatAsString();
        
        int firstRowNum = 0;
        int lastRow = sheet1.getLastRowNum();
        int lastCell = headerList.size() - 1;
        for(int rowNum = 0; rowNum <= lastRow; rowNum++) {
            Row row = sheet1.getRow(rowNum);
            Cell cell = row.getCell(0);
            if(cell == null || cell.toString().equals(Headers.CIKKSZAM.getCat()) || cell.toString().equals("")) {
            	if(cell != null && cell.toString().equals(Headers.CIKKSZAM.getCat()) && !headerList.contains(AddedHeaders.BESZAREUR.getCat())) {
                    if(!headerList.contains(AddedHeaders.BESZAREUR.getCat())) {
                        int i = 1;
                        for(AddedHeaders addedHeaders : AddedHeaders.values()) {
                            Cell addedCell = row.createCell(lastCell + i);
                            addedCell.setCellValue(addedHeaders.getCat());
                            headerList.add(addedCell.getStringCellValue());
                            i++;
                        }
                    }
                    if(cell.toString().equals(Headers.CIKKSZAM.getCat())) {
                        firstRowNum = rowNum;
                    }
                }
                continue;
            }
            
            if(cell.toString().charAt(0) == '3') {
                if(row.getCell(headerList.indexOf(Headers.SZELESSEG2.getCat())) != null) {
                    Cell cellOfWidth = row.getCell(headerList.indexOf(Headers.SZELESSEG2.getCat()));
                    if(!cellOfWidth.toString().equals("")) {
                        double width = 0;
                        width = Double.parseDouble(cellOfWidth.toString());
                        if(width != 0) {
                            Cell cellOfSchemaWidth = row.getCell(headerList.indexOf(Headers.SZELESSEG.getCat()));
                            cellOfSchemaWidth.setCellValue(width / 10 - 100);
                        }
                    }
                }
            }
            
            Cell cellBeszarEur = row.createCell(headerList.indexOf(AddedHeaders.BESZAREUR.getCat()));
            Cell cellBeszarHuf = row.createCell(headerList.indexOf(AddedHeaders.BESZARHUF.getCat()));
            Cell cellAgramBeszar = row.createCell(headerList.indexOf(AddedHeaders.AGRAM_TR.getCat()));
            Cell cellAgramListaar = row.createCell(headerList.indexOf(AddedHeaders.AGRAM_LISTA.getCat()));
            Cell cellOkoviBeszar = row.createCell(headerList.indexOf(AddedHeaders.OKOVI_TR.getCat()));
            Cell cellOkoviListaar = row.createCell(headerList.indexOf(AddedHeaders.OKOVI_LISTA.getCat()));
            Cell cellEurLista = row.createCell(headerList.indexOf(AddedHeaders.EUR_LISTA.getCat()));
            Cell cellHuLista = row.createCell(headerList.indexOf(AddedHeaders.HU_LISTA.getCat()));
            
            Cell cellKatar = row.getCell(headerList.indexOf(Headers.KAT_AR.getCat()));
            Cell cellKatarEng = row.getCell(headerList.indexOf(Headers.RABAT.getCat()));
            Cell cellDeviza = row.getCell(headerList.indexOf(Headers.DEVIZA.getCat()));
            Cell cellFuvar = row.getCell(headerList.indexOf(Headers.FUVAR.getCat()));
            Cell cellVam = row.getCell(headerList.indexOf(Headers.VAM.getCat()));
            Cell cellEngedmeny = row.getCell(headerList.indexOf(Headers.ENGEDMENY.getCat()));
            Cell cellEgyeb = row.getCell(headerList.indexOf(Headers.EGYEB.getCat()));
            Cell cellHulladek = row.getCell(headerList.indexOf(Headers.HULLADEK.getCat()));
            Cell cellSzelesseg = row.getCell(headerList.indexOf(Headers.SZELESSEG.getCat()));
            Cell cellFixktg = row.getCell(headerList.indexOf(Headers.FIXKTG.getCat()));
            
            Cell cellArresEur = row.createCell(headerList.indexOf(AddedHeaders.ARRESEUR.getCat()));
            Cell cellArresHuf = row.createCell(headerList.indexOf(AddedHeaders.ARRESHUF.getCat()));
            Cell cellArresAgram = row.createCell(headerList.indexOf(AddedHeaders.ARRESAGRAM.getCat()));
            Cell cellArresOkovi = row.createCell(headerList.indexOf(AddedHeaders.ARRESOKOVI.getCat()));
            
            String katarPlace = cellKatar.getAddress().formatAsString();
            String katarEngPlace = cellKatarEng.getAddress().formatAsString();
            String fuvarPlace = cellFuvar.getAddress().formatAsString();
            String vamPlace = cellVam.getAddress().formatAsString();
            String engedmenyPlace = cellEngedmeny.getAddress().formatAsString();
            String egyebPlace = cellEgyeb.getAddress().formatAsString();
            String hulladekPlace = cellHulladek.getAddress().formatAsString();
            String szelessegPlace = cellSzelesseg.getAddress().formatAsString();
            String fixktgPlace = cellFixktg.getAddress().formatAsString();
            
            String devizaSzallito = cellDeviza.toString().toUpperCase();
            
            if(devizaSzallito.equals("EUR")) {
                cellBeszarEur.setCellFormula("ROUND(" + katarPlace + "*(1-" + katarEngPlace +"/100)*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1-" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)+" + fixktgPlace + "/" + eurRate2 + ",4)");
                cellBeszarHuf.setCellFormula("ROUND(" + katarPlace + "*(1-" + katarEngPlace +"/100)*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1-" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)*" + eurRate + "+" + fixktgPlace + ",2)");
            }
            else if(devizaSzallito.equals("USD")) {
                cellBeszarEur.setCellFormula("ROUND(" + katarPlace + "*(1-" + katarEngPlace +"/100)*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1-" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)/" + eurUsdRate + "+" + fixktgPlace + "/" + eurRate2 + ",4)");
                cellBeszarHuf.setCellFormula("ROUND(" + katarPlace + "*(1-" + katarEngPlace +"/100)*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1-" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)*" + usdRate + "+" + fixktgPlace + ",2)");
            }
            else if(devizaSzallito.equals("FT") || devizaSzallito.equals("HUF") || devizaSzallito.equals("Ft")) {
                cellBeszarEur.setCellFormula("ROUND((" + katarPlace + "*(1-" + katarEngPlace +"/100)*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1-" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)+" + fixktgPlace + ")/" + eurRate2 + ",4)");
                cellBeszarHuf.setCellFormula("ROUND(" + katarPlace + "*(1-" + katarEngPlace +"/100)*(1+" + fuvarPlace + "/100)*(1+" + vamPlace +
                        "/100)*(1-" + engedmenyPlace + "/100)*(1+" + egyebPlace +"/100)*(1+" + hulladekPlace + 
                        "/100)*(1+" + szelessegPlace + "/100)+" + fixktgPlace + ",2)");
            }
            
            cellAgramBeszar.setCellFormula("ROUND(" + cellBeszarEur.getAddress().formatAsString() + "*1.15,4)");
            cellOkoviBeszar.setCellFormula("ROUND(" + cellBeszarEur.getAddress().formatAsString() + "*1.15,4)");
            cellAgramListaar.setCellFormula("ROUND(" + cellAgramBeszar.getAddress().formatAsString() + "*" +
                    rates[4] + "*" + cellArresAgram.getAddress().formatAsString() + ",2)");
            cellOkoviListaar.setCellFormula("ROUND("+ cellOkoviBeszar.getAddress().formatAsString() + "*" +
                    rates[5] + "*" + cellArresOkovi.getAddress().formatAsString() + ",2)");
            cellEurLista.setCellFormula("ROUND(" + cellBeszarEur.getAddress().formatAsString() + "*" +
                    cellArresEur.getAddress().formatAsString() + "/0.96/0.915,4)");
            cellHuLista.setCellFormula("IF("+ cellBeszarHuf.getAddress().formatAsString() + "*" +
                        cellArresHuf.getAddress().formatAsString() + "<100,ROUND(" + cellBeszarHuf.getAddress().formatAsString() + "*" +
                        cellArresHuf.getAddress().formatAsString() + ",1),ROUND(" + cellBeszarHuf.getAddress().formatAsString() + "*" +
                        cellArresHuf.getAddress().formatAsString() + ",0))");
            
            cellBeszarEur.setCellType(CellType.FORMULA);
            cellBeszarHuf.setCellType(CellType.FORMULA);
            cellAgramBeszar.setCellType(CellType.FORMULA);
            cellOkoviBeszar.setCellType(CellType.FORMULA);
            cellAgramListaar.setCellType(CellType.FORMULA);
            cellOkoviListaar.setCellType(CellType.FORMULA);
            cellEurLista.setCellType(CellType.FORMULA);
            cellHuLista.setCellType(CellType.FORMULA);
            
            CellStyle style1 = wb.createCellStyle();
            style1.setDataFormat(wb.createDataFormat().getFormat("###,###,##0.0000"));
            cellBeszarEur.setCellStyle(style1);
            cellAgramBeszar.setCellStyle(style1);
            cellOkoviBeszar.setCellStyle(style1);
            cellEurLista.setCellStyle(style1);
            
            CellStyle style2 = wb.createCellStyle();
            style2.setDataFormat(wb.createDataFormat().getFormat("###,###,##0.00"));
            cellBeszarHuf.setCellStyle(style2);
            cellAgramListaar.setCellStyle(style2);
            cellOkoviListaar.setCellStyle(style2);
            cellHuLista.setCellStyle(style2);
            cellArresAgram.setCellStyle(style2);
            cellArresEur.setCellStyle(style2);
            cellArresHuf.setCellStyle(style2);
            cellArresOkovi.setCellStyle(style2);
        }

        sheet1.createFreezePane(2, firstRowNum + 1);
        Row row = sheet1.getRow(firstRowNum);
        if(row != null) {
            int lastCellNum = row.getLastCellNum();
            for(int column = 0; column < lastCellNum; column++) {
                sheet1.autoSizeColumn(column);
            }
        }
        
        return wb;
    }
}
