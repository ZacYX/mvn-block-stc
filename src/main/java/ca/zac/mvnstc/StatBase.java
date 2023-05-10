package ca.zac.mvnstc;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

public class StatBase {
    
    static final int HEADER_INDEX = 0;
    static final int CATEGORY_INDEX = 0;
    static final int STOCK_LIST_INDEX = 1;

    ArrayList<StockInfo> stockInfoList;

    Sheet reasonSheet;
    Row currentRow;
    Cell cellWithCategory;
    Cell cellWithStockList;
    String category;    //Reason in updater
    String stockList;
    Boolean oldCategory;

    Integer reasonIndex;
    String sheetName;
    Integer numberOfReasons;

    public StatBase(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet, Integer numberOfReasons) {
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
        return  sheetName;
    }

    void prepare() {
        //Insert a blank column after the first column to the dataSheet, adding 3 to solve outofbounds exception
         reasonSheet.shiftColumns(1, 
             reasonSheet.getRow(ReasonStat.HEADER_INDEX).getLastCellNum() + 3, 1);
        Date date = new Date();
        SimpleDateFormat dateFormatForTitle = new SimpleDateFormat("MMdd");
         reasonSheet.getRow(ReasonStat.HEADER_INDEX).createCell(STOCK_LIST_INDEX).setCellValue(
            dateFormatForTitle.format(date) + " " +  stockInfoList.size());
    }


    void insert() {

    }
}
