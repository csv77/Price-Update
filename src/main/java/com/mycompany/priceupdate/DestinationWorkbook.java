package com.mycompany.priceupdate;

import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DestinationWorkbook {
    private Workbook wb;
    private Sheet sheet;
    private List<List<Cell>> newListOfPrices;
    private DataReorder dataReorder;
    
    public DestinationWorkbook(List<List<Cell>> listOfPrices) throws InvalidFormatException, IOException {
        wb = new XSSFWorkbook();
        sheet = wb.createSheet("PriceUpdate");
        dataReorder = new DataReorder(listOfPrices);
    }

    public Workbook getWb() {
        return wb;
    }

    public void loadData() {
        dataReorder.reorderData();
        newListOfPrices = dataReorder.getOrderedList();
        int i = 0;
        for(List<Cell> cells : newListOfPrices) {
            Row row = sheet.createRow(i++);
            int j = 0;
            for(Cell cell : cells) {
                Cell cellOfNewTable = row.createCell(j++);
                switch(cell.getCellTypeEnum()) {
                    case STRING :
                            cellOfNewTable.setCellValue(cell.getStringCellValue());
                            break;
                        case NUMERIC :
                            cellOfNewTable.setCellValue(cell.getNumericCellValue());
                            break;
                        case BLANK :
                            break;
                        case BOOLEAN :
                            break;
                        case FORMULA :
                            break;
                }
            }
        }
    }
}
