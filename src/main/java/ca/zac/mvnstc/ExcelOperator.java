package ca.zac.mvnstc;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperator {
    private static final String FIRST_REASON_SHEET_NAME = "首因";
    private static final String ALL_REASON_SHEET_NAME = "全因";
    private static final String ALL_REASON_COUNT_SHEET_NAME = "全因数字";

    private FileInputStream updaterFileInputStream;
    private FileInputStream statisticResultFileInputStream;
    private FileOutputStream newStatisticResultFileOutputStream;

    private Workbook updaterWorkbook;
    private Workbook statisticResultWorkbook;

    private Sheet updaterSheet;
    private Sheet firstReasonSheet;
    private Sheet allReasonSheet;
    private Sheet allReasonCountSheet;

    public ExcelOperator(String updaterFilePath, String statisticResultFilePath, String newStatisticResultFilePat) {
        try {
            this.updaterFileInputStream = new FileInputStream(updaterFilePath);
            this.statisticResultFileInputStream = new FileInputStream(statisticResultFilePath);

            Date date = new Date();
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MMdd-hhmm");
            this.newStatisticResultFileOutputStream = new FileOutputStream(
                newStatisticResultFilePat + simpleDateFormat.format(date) + "-marketinfo.xlsx");

            this.updaterWorkbook = new XSSFWorkbook(this.updaterFileInputStream);
            this.statisticResultWorkbook = new XSSFWorkbook(this.statisticResultFileInputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }

        this.updaterSheet = (Sheet) updaterWorkbook.getSheetAt(0); 
        this.firstReasonSheet = (Sheet) statisticResultWorkbook.getSheet(FIRST_REASON_SHEET_NAME);
        this.allReasonSheet = (Sheet) statisticResultWorkbook.getSheet(ALL_REASON_SHEET_NAME);
        this.allReasonCountSheet = (Sheet) statisticResultWorkbook.getSheet(ALL_REASON_COUNT_SHEET_NAME);
    }

    public void close() {
        try {
            this.statisticResultWorkbook.write(newStatisticResultFileOutputStream);
            this.updaterWorkbook.close();
            this.statisticResultWorkbook.close();
            this.updaterFileInputStream.close();
            this.statisticResultFileInputStream.close();
            this.newStatisticResultFileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Sheet getUpdaterSheet() {
        return this.updaterSheet;
    }

    public Sheet getFirstReasonSheet() {
        return this.firstReasonSheet;
    }

    public Sheet getAllReasonSheet() {
        return this.allReasonSheet;
    }

    public Sheet getAllReasonCountSheet() {
        return this.allReasonCountSheet;
    }

    // public Workbook getUpdaterWorkbook() {
    //     return this.updaterWorkbook;
    // }

    // public Workbook getStatisticWorkbook() {
    //     return this.statisticResultWorkbook;
    // }

    // public FileInputStream getUpdaterFileInputStream() {
    //     return this.updaterFileInputStream;
    // }
    
    // public FileInputStream getStatisticResultFileInputStream() {
    //     return this.statisticResultFileInputStream;
    // }

    // public FileOutputStream getNewStatisticResultFileOutputSteam() {
    //     return this.newStatisticResultFileOutputStream;
    // }
}

