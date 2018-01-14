package com.mycompany.priceupdate;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;

public class DataReorder {
    private List<List<Cell>> listOfPrices;
    private List<List<Cell>> orderedList = new ArrayList<>();

    public DataReorder(List<List<Cell>> listOfPrices) {
        this.listOfPrices = listOfPrices;
    }

    public List<List<Cell>> reorderData() {
        int numberOfRows = listOfPrices.size();
        for(int i = 0; i < numberOfRows; i++) {
            int numberOfColumns = listOfPrices.get(i).size();
            for(int j = 1; j < numberOfColumns; j += 2) {
                List<Cell> cells = new ArrayList();
                cells.add(listOfPrices.get(i).get(0));
                cells.add(listOfPrices.get(i).get(j));
                cells.add(listOfPrices.get(i).get(j + 1));
                orderedList.add(cells);
            }
        }
        return orderedList;
    }
}
