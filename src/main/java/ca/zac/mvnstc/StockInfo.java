/**
 * Stock name, increase reason, increase rate and dates in one excel row
 */
package ca.zac.mvnstc;

class StockInfo {
    String name;        //股票名称
    String detailedIndustry;     //细分行业
    String[] reason;      //涨停原因
    Double increaseRate; //涨幅
    Double increaseDates; //连涨天数

    static final String STOCK_NAME_HEADER = "名称";
    static final String STORCK_DETAILED_INDUSTRY_HEADER = "细分行业";
    static final String STOCK_REASON_HEADER = "涨停原因";
    static final String STOCK_INCREASE_RATE_HEADER = "涨幅";
    static final String STOCK_INCREASE_DATES_HEADER = "连涨";
    static final String CELL_EMPTY_STRING = "--";
    static final Double STOCK_INCREASE_FLAG = 0.09;
    static final String STOCK_REASON_SPLITTER_REGEX_STRING = "\\+";
    
    String getName() {
        return this.name;
    }
    String getDetailedIndustry() {
        return this.detailedIndustry;
    }
    String[] getReason() {
        return this.reason;
    }
    Double getIncreaseRate() {
        return this.increaseRate;
    }
    Double getIncreaseDates() {
        return this.increaseDates;
    }
    void setName(String name) {
        this.name = name;
    } 
    void setDetailedIndustry(String detailedIndustry) {
        this.detailedIndustry = detailedIndustry;
    }
    void setReason(String[] reason) {
        this.reason = reason;
    }
    void setIncreaseRate(Double increaseRate) {
        this.increaseRate = increaseRate;
    }
    void setIncreaseDates(Double increaseDates) {
        this.increaseDates = increaseDates;
    }
}
