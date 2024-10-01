/**
 * Insert extracted data to the result excel
 */
package ca.zac.mvnstc;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;

class ReasonCountStat extends StatBase {

    Double stockCount;

    public ReasonCountStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet, Integer numberOfReasons) {
        super(stockInfoList, reasonSheet, numberOfReasons);
    }

    @Override
    void insert() {
        for (int i = 0; i < stockInfoList.size(); i++) {
            if (reasonIndex < stockInfoList.get(i).getReason().length) {
                oldCategory = false; // Assume it is a new item
                // First row is header, iterate from the second row
                for (int j = 1; j <= reasonSheet.getLastRowNum(); j++) {
                    currentRow = reasonSheet.getRow(j);
                    // Get 2 cells
                    cellWithCategory = currentRow.getCell(CATEGORY_INDEX);
                    cellWithStockList = currentRow.getCell(STOCK_LIST_INDEX);
                    if (cellWithCategory != null) {
                        if (cellWithStockList == null) {
                            cellWithStockList = currentRow.createCell(STOCK_LIST_INDEX);
                        }
                    }
                    // Get content of the 2 cells
                    category = cellWithCategory.getStringCellValue().trim();
                    stockCount = cellWithStockList.getNumericCellValue();
                    // Compare reason in arraylist with category in reason statistic excel
                    if (stockInfoList.get(i).getReason()[reasonIndex].equalsIgnoreCase(category)) {
                        stockCount += 1;
                        cellWithStockList.setCellValue(stockCount);
                        oldCategory = true;
                        break; // Category found, do not need to find the rows left
                    }
                }
                // New category, insert a new row
                if (oldCategory == false) {
                    Row newRow = reasonSheet.createRow(reasonSheet.getLastRowNum() + 1);
                    newRow.createCell(CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getReason()[reasonIndex]);
                    newRow.createCell(STOCK_LIST_INDEX).setCellValue(1);
                }
            }
        }

    }

}