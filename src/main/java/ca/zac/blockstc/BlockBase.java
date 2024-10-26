package ca.zac.blockstc;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public abstract class BlockBase {
  /**
   * empty | category | block item
   */
  static final int HEADER_INDEX = 0;
  static final int EMPTY_COLUMN = 0;
  static final int CATEGORY_INDEX = 1;
  static final int BLOCK_DATA_INDEX = 2;

  ArrayList<BlockInfo> blockInfoList;

  Sheet[] outputSheets;
  Row currentRow;
  Cell cellWithCategory;
  Cell cellWithBlockData;
  String category; // Reason in updater
  String blockList;

  public BlockBase(ArrayList<BlockInfo> blockInfoList, Sheet[] outputSheets) {
    this.blockInfoList = blockInfoList;
    this.outputSheets = outputSheets;
  }

  void prepare(Sheet outputSheet) {
    // Write first cell of header for a blank sheet
    if (outputSheet.getLastRowNum() == -1) {
      Row newRow = outputSheet.createRow(BlockBase.HEADER_INDEX);
      newRow.createCell(EMPTY_COLUMN).setCellValue("当日统计");
      newRow.createCell(CATEGORY_INDEX).setCellValue("类别");
    }
    // Insert a blank column after the first column to the dataSheet, adding 3 to
    // solve outofbounds exception
    outputSheet.shiftColumns(BLOCK_DATA_INDEX,
        outputSheet.getRow(BlockBase.HEADER_INDEX).getLastCellNum() + 3, 1);
    Date date = new Date();
    SimpleDateFormat dateFormatForTitle = new SimpleDateFormat("MMdd");
    outputSheet.getRow(BlockBase.HEADER_INDEX).createCell(BLOCK_DATA_INDEX).setCellValue(
        dateFormatForTitle.format(date));
  }

  abstract void insert(Sheet outputSheet, int itermIndex);

  public void process() {
    for (int i = 0; i < outputSheets.length; i++) {
      prepare(outputSheets[i]);
      insert(outputSheets[i], i);
    }
  }
}
