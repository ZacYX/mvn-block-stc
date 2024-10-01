package ca.zac.mvnstc;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;

public class DetailedIndustryCountStat extends StatBase {
    Double stockCount;
    Cell cellWithStockCount;

    public DetailedIndustryCountStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet) {
        super(stockInfoList, reasonSheet, 1);
    }

    @Override
    void insert() {
        // Loop stock array list
        for (int i = 0; i < stockInfoList.size(); i++) {
            // Loop excel
            for (int j = 1; j <= reasonSheet.getLastRowNum() + 1; j++) {
                // This is a blank row, a new row has to be create
                if (j == reasonSheet.getLastRowNum() + 1) {
                    Row newRow = reasonSheet.createRow(j);
                    // Write category name
                    newRow.createCell(CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getDetailedIndustry());
                    // write stock count, new row must be 1
                    newRow.createCell(STOCK_LIST_INDEX).setCellValue(1);
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
                cellWithStockCount = currentRow.getCell(STOCK_LIST_INDEX);
                if (cellWithStockCount == null) {
                    cellWithStockCount = currentRow.createCell(STOCK_LIST_INDEX);
                }
                stockCount = cellWithStockCount.getNumericCellValue() + 1;
                cellWithStockCount.setCellValue(stockCount);
                // Don't need to compare the following rows
                break;
            }

        }
    }
}
