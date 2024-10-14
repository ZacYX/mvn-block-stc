/**
 * Insert extracted data to the result excel
 */
package ca.zac.mvnstc;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;

class ReasonStat extends StatBase {

    public ReasonStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet, Integer numberOfReasons) {
        super(stockInfoList, reasonSheet, numberOfReasons);
    }

    @Override
    void insert() {
        // Loop excel sheet rows to set second column cells 0
        if (reasonIndex < 1) {
            for (int j = 1; j <= reasonSheet.getLastRowNum(); j++) {
                currentRow = reasonSheet.getRow(j);
                if (currentRow == null) {
                    continue;
                }
                if (currentRow.getCell(SECOND_COLUMN) == null) {
                    continue;
                }
                currentRow.getCell(SECOND_COLUMN).setCellValue(0);
            }
        }
        // Loop stock list
        for (int i = 0; i < stockInfoList.size(); i++) {
            if (reasonIndex < stockInfoList.get(i).getReason().length) {
                oldCategory = false; // Assume it is a new item
                // First row is header, iterate from the second row
                // Loop excell rows
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
                    stockList = cellWithStockList.getStringCellValue();
                    // Compare reason in arraylist with category in reason statistic excel
                    if (stockInfoList.get(i).getReason()[reasonIndex].equalsIgnoreCase(category)) {
                        // Write increase dates that is greater than 1 at the end of each stock name
                        if (stockInfoList.get(i).getIncreaseDates() > 1) {
                            stockList += stockInfoList.get(i).getName()
                                    + stockInfoList.get(i).getIncreaseDates().intValue() + "\n";
                        } else {
                            stockList += stockInfoList.get(i).getName() + "\n";
                        }
                        cellWithStockList.setCellValue(stockList);
                        // update number of second column
                        currentRow.getCell(SECOND_COLUMN)
                                .setCellValue(currentRow.getCell(SECOND_COLUMN).getNumericCellValue() + 1);
                        oldCategory = true;
                        break; // Category found, do not need to find the rows left
                    }
                }
                // New category, insert a new row
                if (oldCategory == false) {
                    Row newRow = reasonSheet.createRow(reasonSheet.getLastRowNum() + 1);
                    newRow.createCell(CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getReason()[reasonIndex]);
                    if (stockInfoList.get(i).getIncreaseDates() > 1) {
                        newRow.createCell(STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName()
                                + stockInfoList.get(i).getIncreaseDates().intValue() + "\n");
                    } else {
                        newRow.createCell(STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName() + "\n");
                    }
                    // set second column with 1
                    newRow.createCell(SECOND_COLUMN).setCellValue(1);
                }
            }
        }

    }

}