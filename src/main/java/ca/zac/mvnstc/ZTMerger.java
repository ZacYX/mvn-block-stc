/**
 * cmd example: java ca.stc.merger.ZTMerger "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx" "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx" "C:\\Users\\User\\Documents\\stcdata\\"
 */
package ca.zac.mvnstc;

import java.util.ArrayList;

public class ZTMerger {
   static String marketInfoPath = "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx";
   static String updaterPath = "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx";
   static String outputPath = "C:\\Users\\User\\Documents\\stcdata\\";

   public static void main(String args[]) {
      if (args.length == 3) {
         marketInfoPath = args[0];
         updaterPath = args[1];
         outputPath = args[2];
      }
      ExcelOperator excelOperator = null;
      try {
         excelOperator = new ExcelOperator(updaterPath, marketInfoPath, outputPath);

         Updater updater = new Updater(excelOperator.getUpdaterSheet());
         ArrayList<StockInfo> stockInfoList = updater.getData();

         ReasonStat reasonStat = new ReasonStat(stockInfoList, excelOperator.getFirstReasonSheet(), 1);
         reasonStat.process();
         ReasonStat allReasonStat = new ReasonStat(stockInfoList, excelOperator.getAllReasonSheet(), 4);
         allReasonStat.process();
         ReasonCountStat allReasonCountStat = new ReasonCountStat(stockInfoList, excelOperator.getAllReasonCountSheet(),
               4);
         allReasonCountStat.process();
         DetailedIndustryStat detailedIndustryStat = new DetailedIndustryStat(stockInfoList,
               excelOperator.getDetailedIndustrySheet());
         detailedIndustryStat.process();
         DetailedIndustryCountStat detailedIndustryCountStat = new DetailedIndustryCountStat(stockInfoList,
               excelOperator.getDetailedIndustryCountSheet());
         detailedIndustryCountStat.process();

      } catch (Exception e) {
         System.out.println("Excetion in main");
         e.printStackTrace();
      } finally {
         System.out.println("Finally in main");
         excelOperator.close();
      }

   }
}