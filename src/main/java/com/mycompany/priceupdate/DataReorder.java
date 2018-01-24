package com.mycompany.priceupdate;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;

public class DataReorder {
    private List<List<Cell>> listOfData;
    private List<List<Cell>> orderedList = new ArrayList<>();

    public DataReorder(List<List<Cell>> listOfData) {
        this.listOfData = listOfData;
    }

    public List<List<Cell>> reorderData() {
        int numberOfRows = listOfData.size();
        for(int i = 0; i < numberOfRows; i++) {
            int numberOfColumns = listOfData.get(i).size();
            for(int j = 1; j < numberOfColumns; j += 3) {
                List<Cell> cells = new ArrayList<Cell>();
                cells.add(listOfData.get(i).get(0));
                cells.add(listOfData.get(i).get(j));
                cells.add(listOfData.get(i).get(j + 1));
                cells.add(listOfData.get(i).get(j + 2));
                orderedList.add(cells);
            }
        }
        return orderedList;
    }
}
