package ca.zac.blockstc;

public class TableHead {
  String name;
  int index;

  public TableHead(String name, int index) {
    this.name = name;
    this.index = index;
  }

  public String getName() {
    return this.name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public int getIndex() {
    return this.index;
  }

  public void setIndex(int index) {
    this.index = index;
  }

}
