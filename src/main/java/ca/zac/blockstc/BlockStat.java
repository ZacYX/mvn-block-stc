
package ca.zac.blockstc;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class BlockStat extends BlockBase {

  public BlockStat(ArrayList<BlockInfo> blockInfoList, Sheet[] outputSheet) {
    super(blockInfoList, outputSheet);
  }

  @Override
  void insert(Sheet outputSheet, int itemIndex) {
    // Loop stock array list
    for (int i = 0; i < blockInfoList.size(); i++) {
      // Loop excel
      for (int j = 1; j <= outputSheet.getLastRowNum() + 1; j++) {
        // This is a blank row, a new row has to be create
        if (j == outputSheet.getLastRowNum() + 1) {
          Row newRow = outputSheet.createRow(j);
          // Write category name
          newRow.createCell(CATEGORY_INDEX).setCellValue(blockInfoList.get(i).getTitle());
          // write stock name with dates
          if (blockInfoList.get(i).getItemData()[itemIndex] instanceof String) {
            newRow.createCell(BLOCK_DATA_INDEX).setCellValue((String) blockInfoList.get(i).getItemData()[itemIndex]);
          }
          if (blockInfoList.get(i).getItemData()[itemIndex] instanceof Double) {
            newRow.createCell(BLOCK_DATA_INDEX).setCellValue((Double) blockInfoList.get(i).getItemData()[itemIndex]);
          }
          break;
        }
        // Compare existing category
        currentRow = outputSheet.getRow(j);
        // blank row
        if (currentRow == null) {
          continue;
        }
        cellWithCategory = currentRow.getCell(CATEGORY_INDEX);
        // Category can't be null or not the same category
        if (cellWithCategory == null
            || !cellWithCategory.getStringCellValue().trim().equalsIgnoreCase(
                blockInfoList.get(i).getTitle())) {
          continue;
        }
        // Found existing category
        cellWithBlockData = currentRow.getCell(BLOCK_DATA_INDEX);
        if (cellWithBlockData == null) {
          cellWithBlockData = currentRow.createCell(BLOCK_DATA_INDEX);
        }
        if (blockInfoList.get(i).getItemData()[itemIndex] instanceof String) {
          cellWithBlockData.setCellValue((String) blockInfoList.get(i).getItemData()[itemIndex]);
        }
        if (blockInfoList.get(i).getItemData()[itemIndex] instanceof Double) {
          cellWithBlockData.setCellValue((Double) blockInfoList.get(i).getItemData()[itemIndex]);
        }
        // Don't need to compare the following rows
        break;
      }

    }
  }

}