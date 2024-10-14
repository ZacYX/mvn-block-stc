package ca.zac.mvnstc;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class DetailedIndustryStat extends StatBase {

    public DetailedIndustryStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet) {
        super(stockInfoList, reasonSheet, 1); // One stock has only one detailed industry
    }

    @Override
    void insert() {
        // Loop excel and set the second column cells 0
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
        // Loop stock array list
        for (int i = 0; i < stockInfoList.size(); i++) {
            // Loop excel
            for (int j = 1; j <= reasonSheet.getLastRowNum() + 1; j++) {
                // This is a blank row, a new row has to be create
                if (j == reasonSheet.getLastRowNum() + 1) {
                    Row newRow = reasonSheet.createRow(j);
                    // Write category name
                    newRow.createCell(CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getDetailedIndustry());
                    // write stock name with dates
                    if (stockInfoList.get(i).getIncreaseDates() > 1) {
                        newRow.createCell(STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName()
                                + stockInfoList.get(i).getIncreaseDates().intValue() + "\n");
                    } else {
                        newRow.createCell(STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName() + "\n");
                    }
                    // write the second column with 1
                    newRow.createCell(SECOND_COLUMN).setCellValue(1);
                    break;
                }
                // Compare existing category
                currentRow = reasonSheet.getRow(j);
                // blank row
                if (currentRow == null) {
                    continue;
                }
                cellWithCategory = currentRow.getCell(CATEGORY_INDEX);
                // Category can't be null
                if (cellWithCategory == null
                        || !cellWithCategory.getStringCellValue().trim().equalsIgnoreCase(
                                stockInfoList.get(i).getDetailedIndustry())) {
                    continue;
                }
                // Found existing category
                cellWithStockList = currentRow.getCell(STOCK_LIST_INDEX);
                if (cellWithStockList == null) {
                    cellWithStockList = currentRow.createCell(STOCK_LIST_INDEX);
                }
                stockList = cellWithStockList.getStringCellValue() + stockInfoList.get(i).getName();
                if (stockInfoList.get(i).getIncreaseDates() > 1) {
                    stockList += stockInfoList.get(i).getIncreaseDates().intValue();
                }
                stockList += "\n";
                cellWithStockList.setCellValue(stockList);
                // update number of the second colunm
                currentRow.getCell(SECOND_COLUMN)
                        .setCellValue(currentRow.getCell(SECOND_COLUMN).getNumericCellValue() + 1);
                // Don't need to compare the following rows
                break;
            }

        }
    }

}