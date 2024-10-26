package ca.zac.blockstc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperator {

    // default value of input and output files
    static String newResultFilePath = "C:\\Users\\User\\Documents\\stcdata\\";
    static String resultToChangeFilePath = "C:\\Users\\User\\Documents\\stcdata\\industry.xlsx";
    static String updaterFilePath = "C:\\Users\\User\\Documents\\stcdata\\u.xlsx";

    String outputFileName = "block-info";

    FileOutputStream newResultFileOutputStream;
    FileInputStream resultToChangeFileInputStream;
    FileInputStream updaterFileInputStream;

    // input workbook and sheet
    Workbook updaterWorkbook;
    Sheet updaterSheet;

    Workbook resultToChangeWorkbook;
    Sheet[] itemSheetsToChange;

    public ExcelOperator() throws IOException {
        // input file updater with raw information
        // do this in constructor to make sure BlockInfo.items.size exist
        // before getOuptutSheets
        try {
            this.updaterFileInputStream = new FileInputStream(updaterFilePath);
            updaterWorkbook = new XSSFWorkbook(this.updaterFileInputStream);
            updaterSheet = (Sheet) updaterWorkbook.getSheetAt(0);
        } catch (IOException e) {
            System.out.println("Open updater failed!");
            throw e;
        }
    }

    // Input sheets
    public Sheet getUpdaterSheet() {
        return this.updaterSheet;
    }

    // block output sheets
    public Sheet[] getOutputSheets() {
        if (BlockInfo.items.isEmpty()) {
            System.err.println("Must find items in updater first");
            return null;
        }
        itemSheetsToChange = new Sheet[BlockInfo.items.size()];
        // input file with previously result data
        try {
            File file = new File(updaterFilePath);
            String fileName = file.getName();
            outputFileName = fileName.contains(".")
                    ? fileName.substring(0, fileName.lastIndexOf("."))
                    : fileName;
            resultToChangeFileInputStream = new FileInputStream(resultToChangeFilePath);
            resultToChangeWorkbook = new XSSFWorkbook(resultToChangeFileInputStream);
        } catch (IOException e) {
            System.out.println("marketinfo file not found, create a new one");
            resultToChangeWorkbook = new XSSFWorkbook();
        }
        // read every sheet specified in args
        try {
            Date date = new Date();
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MMdd-hhmm");
            newResultFileOutputStream = new FileOutputStream(
                    newResultFilePath + simpleDateFormat.format(date) + "-" + outputFileName + ".xlsx");

            // Block result output sheets
            for (int i = 0; i < BlockInfo.items.size(); i++) {
                itemSheetsToChange[i] = resultToChangeWorkbook.getSheet(BlockInfo.items.get(i).getName());
                if (itemSheetsToChange[i] == null) {
                    itemSheetsToChange[i] = resultToChangeWorkbook.createSheet(BlockInfo.items.get(i).getName());
                }
            }
        } catch (Exception e) {
            System.out.println("Exception in ExcelOperator");
            e.printStackTrace();
        }
        return itemSheetsToChange;
    }

    public void close() {
        try {
            resultToChangeWorkbook.write(newResultFileOutputStream);
            updaterWorkbook.close();
            resultToChangeWorkbook.close();
            updaterFileInputStream.close();
            resultToChangeFileInputStream.close();
            newResultFileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
