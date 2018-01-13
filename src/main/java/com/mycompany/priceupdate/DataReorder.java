package com.mycompany.priceupdate;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;

public class DataReorder {
    private List<List<Cell>> listOfPrices;
    private List<List<Cell>> orderedList = new ArrayList<>();
    private int numberOfRows;
    private int numberOfColumns;

    public DataReorder(List<List<Cell>> listOfPrices) {
        this.listOfPrices = listOfPrices;
        numberOfRows = listOfPrices.size();
        numberOfColumns = listOfPrices.get(0).size();
    }

    public List<List<Cell>> getOrderedList() {
        return orderedList;
    }
    
    public List<List<Cell>> reorderData() {
        for(int i = 0; i < numberOfRows; i++) {
            for(int j = 1; j < numberOfColumns; j++) {
                List<Cell> cells = new ArrayList();
                cells.add(listOfPrices.get(i).get(0));
                cells.add(listOfPrices.get(i).get(j));
                orderedList.add(cells);
            }
        }
        
        return orderedList;
    }
}
