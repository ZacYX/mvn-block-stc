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
    static final int SECOND_COLUMN = 1;
    static final int STOCK_LIST_INDEX = 2;

    ArrayList<StockInfo> stockInfoList;

    Sheet reasonSheet;
    Row currentRow;
    Cell cellWithCategory;
    Cell cellWithStockList;
    String category; // Reason in updater
    String stockList;
    Boolean oldCategory;

    Integer reasonIndex; // Concept reason index
    String sheetName;
    Integer numberOfReasons; // Many 10% increased stocks has more than one concept reason

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
        return sheetName;
    }

    void prepare() {
        // Write first cell of header for a blank sheet
        if (reasonSheet.getLastRowNum() == -1) {
            Row newRow = reasonSheet.createRow(StatBase.HEADER_INDEX);
            newRow.createCell(CATEGORY_INDEX).setCellValue("类别");
            newRow.createCell(SECOND_COLUMN).setCellValue("当日统计");
        }
        // Insert a blank column after the first column to the dataSheet, adding 3 to
        // solve outofbounds exception
        reasonSheet.shiftColumns(STOCK_LIST_INDEX,
                reasonSheet.getRow(StatBase.HEADER_INDEX).getLastCellNum() + 3, 1);
        Date date = new Date();
        SimpleDateFormat dateFormatForTitle = new SimpleDateFormat("MMdd");
        reasonSheet.getRow(StatBase.HEADER_INDEX).createCell(STOCK_LIST_INDEX).setCellValue(
                dateFormatForTitle.format(date) + " " + stockInfoList.size());
    }

    void insert() {

    }
}
