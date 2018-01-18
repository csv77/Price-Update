package com.mycompany.priceupdate;

import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
        sheet = wb.createSheet("PricesToUpload");
        dataReorder = new DataReorder(listOfPrices);
    }

    public Workbook getWb() {
        return wb;
    }

    public void loadData() {
        newListOfPrices = dataReorder.reorderData();
        
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
                        if(j == 2) {
                            CellStyle style = wb.createCellStyle();
                            style.setDataFormat(wb.createDataFormat().getFormat("###,###,##0.0000"));
                            cellOfNewTable.setCellStyle(style);
                        }
                        break;
                    case BLANK :
                        break;
                    case _NONE :
                        break;
                    case FORMULA :
                        cellOfNewTable.setCellFormula(cell.getCellFormula());
                        break;
                }
            }
        }
        int lastCellNum = sheet.getRow(0).getLastCellNum();
        for(int column = 0; column < lastCellNum; column++) {
            sheet.autoSizeColumn(column);
        }
    }
}
