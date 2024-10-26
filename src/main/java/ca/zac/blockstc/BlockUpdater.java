package ca.zac.blockstc;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.*;

public class BlockUpdater {
  Sheet updaterSheet;
  Row currentRow;
  Cell cellWithName;
  Cell[] cellWithItems;

  BlockInfo blockInfo;
  ArrayList<BlockInfo> blockInfoList;

  public BlockUpdater(Sheet updaterSheet) {
    this.updaterSheet = updaterSheet;
  }

  public ArrayList<BlockInfo> getData() {
    this.prepare();
    this.process();
    return this.blockInfoList;
  }

  // prepare workbook, worksheet, collumn index of name, leadStock, increase rate
  // and count
  void prepare() {
    // BlockInfo.itemColumnIndexes = new int[BlockInfo.items.size()];
    Row headerRow = updaterSheet.getRow(0); // First row
    for (Cell cell : headerRow) {
      // There is a space before this string
      if (cell.getStringCellValue().trim().contains(BlockInfo.blockTitle.getName())) {
        BlockInfo.blockTitle.setIndex(cell.getColumnIndex());
      }
      for (int i = 0; i < BlockInfo.itemsInput.size(); i++) {
        if (cell.getStringCellValue().trim().equals(BlockInfo.itemsInput.get(i))) {
          TableHead tableHead = new TableHead(BlockInfo.itemsInput.get(i), cell.getColumnIndex());
          BlockInfo.items.add(tableHead);
        }
      }
    }
    System.out.println(BlockInfo.blockTitle.getName() + " Index: " + BlockInfo.blockTitle.getIndex());
    for (int i = 0; i < BlockInfo.items.size(); i++) {
      System.out.println(BlockInfo.items.get(i).getName() + " index: " + BlockInfo.items.get(i).getIndex());
    }
  }

  // Get block name, lead stock, increase rate, increase count according row index
  void process() {
    cellWithItems = new Cell[BlockInfo.items.size()];
    blockInfoList = new ArrayList<BlockInfo>();
    blockInfo = new BlockInfo();
    for (int i = 1; i <= updaterSheet.getLastRowNum(); i++) {
      // get cells in a row of excel
      currentRow = updaterSheet.getRow(i);
      cellWithName = currentRow.getCell(BlockInfo.blockTitle.getIndex());
      for (int j = 0; j < BlockInfo.items.size(); j++) {
        cellWithItems[j] = currentRow.getCell(BlockInfo.items.get(j).getIndex());
      }
      // read cells' content
      try {
        // Read Name from cell and store it
        blockInfo.setTitle(cellWithName.getStringCellValue().trim());
        // Read items' data from cells and store them
        for (int j = 0; j < BlockInfo.items.size(); j++) {
          if (cellWithItems[j].getCellType() == CellType.STRING) {
            blockInfo.setItemData(j, cellWithItems[j].getStringCellValue().trim());
          }
          if (cellWithItems[j].getCellType() == CellType.NUMERIC) {
            blockInfo.setItemData(j, cellWithItems[j].getNumericCellValue());
          }
        }

        // Not "--" and increase rate > 0.09 means a valid info, and add it to the list
        if (blockInfo.getTitle().length() > 0) {
          blockInfoList.add(this.blockInfo);
          blockInfo = new BlockInfo();
        }
      } catch (Exception e) {
        e.printStackTrace();
        continue;
      }
    }
    System.out.println("Total block items: " + blockInfoList.size());
  }
}
