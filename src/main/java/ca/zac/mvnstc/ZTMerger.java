/**
 * cmd example: java ca.stc.merger.ZTMerger "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx" "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx" "C:\\Users\\User\\Documents\\stcdata\\"
 */
package ca.zac.mvnstc;

public class ZTMerger {
      static String marketInfoPath = "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx";
      static String updaterPath = "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx";
      static String outputPath = "C:\\Users\\User\\Documents\\stcdata\\";
   public static void main(String args[]) {
      if(args.length == 3) {
         marketInfoPath = args[0];
         updaterPath = args[1];
         outputPath = args[2];
      }
      
      ExcelOperator excelOperator = new ExcelOperator(updaterPath, marketInfoPath, outputPath);

      Updater updater = new Updater(excelOperator.getUpdaterSheet());

      ReasonStat reasonStat = new ReasonStat(updater.getData(), excelOperator.getFirstReasonSheet(), 1);
      reasonStat.process();
      ReasonStat allReasonStat = new ReasonStat(updater.getData(), excelOperator.getAllReasonSheet(), 4);
      allReasonStat.process();
      ReasonCountStat allReasonCountStat = new ReasonCountStat(updater.getData(), excelOperator.getAllReasonCountSheet(), 4);
      allReasonCountStat.process();

      excelOperator.close();

   }  
}