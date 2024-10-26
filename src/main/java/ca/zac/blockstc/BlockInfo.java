package ca.zac.blockstc;

import java.util.ArrayList;

public class BlockInfo {

  static final String BLOCK_NAME_HEADER = "板块名称";
  static final String CELL_EMPTY_STRING = "--";

  static ArrayList<String> itemsInput = new ArrayList<>(); // 涨幅， 涨停数， 领涨股， 10日涨幅...

  // column index in excel
  static TableHead blockTitle = new TableHead(BLOCK_NAME_HEADER, -1);
  static ArrayList<TableHead> items = new ArrayList<>();

  String title; // 板块名称
  Object[] itemData;

  public BlockInfo() {
    itemData = new Object[items.size()];
  }

  public String getTitle() {
    return this.title;
  }

  public void setTitle(String title) {
    this.title = title;
  }

  public Object[] getItemData() {
    return this.itemData;
  }

  public void setItemData(int index, Object data) {
    this.itemData[index] = data;
  }
}
