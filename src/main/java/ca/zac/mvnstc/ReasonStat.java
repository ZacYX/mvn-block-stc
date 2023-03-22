/**
 * Insert extracted data to the result excel
 */
package ca.zac.mvnstc;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;

class ReasonStat {
    private static final int HEADER_INDEX = 0;
    private static final int CATEGORY_INDEX = 0;
    private static final int STOCK_LIST_INDEX = 1;

    private ArrayList<StockInfo> stockInfoList;

    private Sheet reasonSheet;
    private Row currentRow;
    private Cell cellWithCategory;
    private Cell cellWithStockList;
    private String category;    //Reason in updater
    private String stockList;
    private Boolean oldCategory;

    private Integer reasonIndex;
    private String sheetName;
    private Integer numberOfReasons;

    public ReasonStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet, Integer numberOfReasons) {
        this.stockInfoList = stockInfoList;
        this.reasonSheet = reasonSheet;
        this.numberOfReasons = numberOfReasons;
    }

    public void process() {
        prepare();
        for (int i = 0; i < numberOfReasons; i++) {
            setReasonIndex(i);
            insert();
        }
    }

    public void setReasonIndex(Integer reasonIndex) {
        this.reasonIndex = reasonIndex;
    }
    public Integer getReasonIndex() {
        return this.reasonIndex;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
    public String getSheetName() {
        return this.sheetName;
    }

    void prepare() {
        //Insert a blank column after the first column to the dataSheet, adding 3 to solve outofbounds exception
        this.reasonSheet.shiftColumns(1, 
            this.reasonSheet.getRow(ReasonStat.HEADER_INDEX).getLastCellNum() + 3, 1);
        Date date = new Date();
        SimpleDateFormat dateFormatForTitle = new SimpleDateFormat("MMdd");
        this.reasonSheet.getRow(ReasonStat.HEADER_INDEX).createCell(STOCK_LIST_INDEX).setCellValue(
            dateFormatForTitle.format(date) + " " + this.stockInfoList.size());
    }

    void insert() {
        for (int i = 0; i < this.stockInfoList.size(); i++) {
            if (this.reasonIndex < stockInfoList.get(i).getReason().length) {   
                this.oldCategory = false;       //Assume it is a new item
                //First row is header, iterate from the second row
                for (int j = 1; j <= this.reasonSheet.getLastRowNum(); j++) {
                    this.currentRow = this.reasonSheet.getRow(j);
                    //Get 2 cells
                    this.cellWithCategory = this.currentRow.getCell(ReasonStat.CATEGORY_INDEX); 
                    this.cellWithStockList = this.currentRow.getCell(ReasonStat.STOCK_LIST_INDEX);
                    if (this.cellWithCategory != null) {
                        if (this.cellWithStockList == null) {
                            this.cellWithStockList = this.currentRow.createCell(ReasonStat.STOCK_LIST_INDEX);
                        }
                    }
                    //Get content of the 2 cells
                    this.category = this.cellWithCategory.getStringCellValue().trim();
                    this.stockList = this.cellWithStockList.getStringCellValue();
                    //Compare reason in arraylist with category in reason statistic excel
                    if (stockInfoList.get(i).getReason()[this.reasonIndex].equalsIgnoreCase(this.category)) {
                        //Write increase dates that is greater than 1 at the end of each stock name
                        if(stockInfoList.get(i).getIncreaseDates() > 1) {
                            this.stockList += stockInfoList.get(i).getName() 
                                + stockInfoList.get(i).getIncreaseDates().intValue() + "\n";
                        } else {
                            this.stockList += stockInfoList.get(i).getName() + "\n";
                        }
                        this.cellWithStockList.setCellValue(this.stockList);
                        this.oldCategory = true;
                        break;   //Category found, do not need to find the rows left
                    }
                }
                //New category, insert a new row
                if (this.oldCategory == false) {
                    Row newRow = this.reasonSheet.createRow(this.reasonSheet.getLastRowNum() + 1);
                    newRow.createCell(ReasonStat.CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getReason()[this.reasonIndex]);
                    if (stockInfoList.get(i).getIncreaseDates() > 1) {
                        newRow.createCell(ReasonStat.STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName() 
                            + stockInfoList.get(i).getIncreaseDates().intValue() + "\n"); 
                    } else {
                        newRow.createCell(ReasonStat.STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName() + "\n"); 
                    }
                }
            }
        }
        
    }


}