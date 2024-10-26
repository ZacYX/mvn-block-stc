/**
 *chcp 65001    //this line is to change cmd line encode utf-8
 *java  -jar "C:\\Users\\User\\OneDrive\\StockData\\block-stc-1.0.0.jar" "C:\\Users\\User\\Documents\\stcdata\\" "C:\\Users\\User\\Documents\\stcdata\\in.xlsx" "C:\\Users\\User\\Documents\\stcdata\\inu.xlsx" "涨停数" "领涨股" "涨幅" "5日涨幅" "10日涨幅" "20日涨幅"
 *echo merge done! 
 *pause 
 */
package ca.zac.blockstc;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Sheet;

public class BlockMerger {

   public static void main(String args[]) {
      if (args.length < 3) {
         System.err.println("At least 3 args ");
         return;
      }
      if (args.length >= 3) {
         ExcelOperator.newResultFilePath = args[0];
         ExcelOperator.resultToChangeFilePath = args[1];
         ExcelOperator.updaterFilePath = args[2];

         System.out.println("output path: " + ExcelOperator.newResultFilePath);
         System.out.println("resul to change path: " + ExcelOperator.resultToChangeFilePath);
         System.out.println("updater path: " + ExcelOperator.updaterFilePath);
      }
      if (args.length > 3) {
         for (int i = 0; i < args.length - 3; i++) {
            BlockInfo.itemsInput.add(args[i + 3]);
            System.out.println("statItems " + i + " is: " + BlockInfo.itemsInput.get(i));
         }
      }
      ExcelOperator excelOperator = null;
      try {
         excelOperator = new ExcelOperator();

         BlockUpdater blockUpdater = new BlockUpdater(excelOperator.getUpdaterSheet());
         ArrayList<BlockInfo> blockInfoList = blockUpdater.getData();

         // Must behind blockupdater.getData, BlockInfo.items can not be empty
         Sheet[] outputSheets = excelOperator.getOutputSheets();

         BlockStat blockStat = new BlockStat(blockInfoList, outputSheets);
         blockStat.process();

      } catch (Exception e) {
         System.out.println("Excetion in main");
         e.printStackTrace();
      } finally {
         System.out.println("Finally in main");
         excelOperator.close();
      }

   }
}