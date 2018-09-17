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
    private Sheet sheet1;
    private Sheet sheet2;
    private List<List<Cell>> newListOfPrices;
    private DataReorder pricesReorder;
    private List<List<Cell>> newListOfSchema;
    private DataReorder schemaReorder;
    
    public DestinationWorkbook(List<List<Cell>> listOfPrices,
            List<List<Cell>> listOfSchema) throws InvalidFormatException, IOException {
        wb = new XSSFWorkbook();
        sheet1 = wb.createSheet("PricesToUpload");
        pricesReorder = new DataReorder(listOfPrices);
        sheet2 = wb.createSheet("SchemaToUpload");
        schemaReorder = new DataReorder(listOfSchema);
    }

    public Workbook getWb() {
        loadData(newListOfPrices, pricesReorder, sheet1);
        loadData(newListOfSchema, schemaReorder, sheet2);
        return wb;
    }

    @SuppressWarnings("incomplete-switch")
	private void loadData(List<List<Cell>> newListOfData, DataReorder dataReorder,
            Sheet sheet) {
        newListOfData = dataReorder.reorderData();
        
        int i = 0;
        for(List<Cell> cells : newListOfData) {
            Row row = sheet.createRow(i++);
            int j = 0;
            for(Cell cell : cells) {
                Cell cellOfNewTable = row.createCell(j++);
                switch(cell.getCellTypeEnum()) {
                    case STRING :
                        cellOfNewTable.setCellValue(cell.getStringCellValue());
                        break;
                    case NUMERIC : case FORMULA :
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
                }
            }
        }
        Row row = sheet.getRow(0);
        if(row != null) {
            int lastCellNum = row.getLastCellNum();
            for(int column = 0; column < lastCellNum; column++) {
                sheet.autoSizeColumn(column);
            }
        }
    }
}
