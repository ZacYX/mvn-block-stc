/**
 * cmd example: java ca.stc.merger.ZTMerger "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx" "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx" "C:\\Users\\User\\Documents\\stcdata\\"
 */
package ca.zac.mvnstc;

public class ZTMerger {
   public static void main(String args[]) {
      String marketInfoPath = args[0];
      String updaterPath = args[1];
      String outputPath = args[2];
      // String marketInfoPath = "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx";
      // String updaterPath = "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx";
      // String outputPath = "C:\\Users\\User\\Documents\\stcdata\\";
      
      ExcelOperator excelOperator = new ExcelOperator(updaterPath, marketInfoPath, outputPath);

      Updater updater = new Updater(excelOperator.getUpdaterSheet());

      ReasonStat reasonStat = new ReasonStat(updater.getData(), excelOperator.getFirstReasonSheet(), 1);
      reasonStat.process();
      ReasonStat allReasonStat = new ReasonStat(updater.getData(), excelOperator.getAllReasonSheet(), 4);
      allReasonStat.process();
      
      excelOperator.close();

   }  
}