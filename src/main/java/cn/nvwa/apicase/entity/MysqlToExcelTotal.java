package cn.nvwa.apicase.entity;

/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class MysqlToExcelTotal {
    private static final long serialVersionUID=1L;

    private int total_tesecase_num;//总用例数
    private String bug_total_percent;//bug总占比
    private String total_bug_type;//bug类型
    private int total_bug_type_bum;//bug类型统计总数
    private String total_percent;//bug类型总占比

    public MysqlToExcelTotal() {
    }

    public int getTotal_tesecase_num() {
        return total_tesecase_num;
    }

    public void setTotal_tesecase_num(int total_tesecase_num) {
        this.total_tesecase_num = total_tesecase_num;
    }

    public String getTotal_bug_type() {
        return total_bug_type;
    }

    public void setTotal_bug_type(String total_bug_type) {
        this.total_bug_type = total_bug_type;
    }

    public int getTotal_bug_type_bum() {
        return total_bug_type_bum;
    }

    public void setTotal_bug_type_bum(int total_bug_type_bum) {
        this.total_bug_type_bum = total_bug_type_bum;
    }

    public String getTotal_percent() {
        return total_percent;
    }

    public void setTotal_percent(String total_percent) {
        this.total_percent = total_percent;
    }

    public String getBug_total_percent() {
        return bug_total_percent;
    }

    public void setBug_total_percent(String bug_total_percent) {
        this.bug_total_percent = bug_total_percent;
    }
}
