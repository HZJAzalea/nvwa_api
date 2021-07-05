package cn.nvwa.apicase.entity;

/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class MysqlToExcel {
    private static final long serialVersionUID=1L;

    private String interface_name; //接口名称
    private String interface_url; //接口路径
    private int testcase_num; //总用例数
    private int total_bug_num;//bug总数
    private String total_bug_percent;//bug总占比
    private String bug_type; //bug类型
    private int bug_type_num; //bug类型统计数
    private String percent; //bug类型百分比

    public MysqlToExcel() {
    }


    public String getInterface_name() {
        return interface_name;
    }

    public void setInterface_name(String interface_name) {
        this.interface_name = interface_name;
    }

    public String getInterface_url() {
        return interface_url;
    }

    public void setInterface_url(String interface_url) {
        this.interface_url = interface_url;
    }

    public int getTestcase_num() {
        return testcase_num;
    }

    public void setTestcase_num(int testcase_num) {
        this.testcase_num = testcase_num;
    }

    public String getBug_type() {
        return bug_type;
    }

    public void setBug_type(String bug_type) {
        this.bug_type = bug_type;
    }

    public int getBug_type_num() {
        return bug_type_num;
    }

    public void setBug_type_num(int bug_type_num) {
        this.bug_type_num = bug_type_num;
    }

    public String getPercent() {
        return percent;
    }

    public void setPercent(String percent) {
        this.percent = percent;
    }

    public int getTotal_bug_num() {
        return total_bug_num;
    }

    public void setTotal_bug_num(int total_bug_num) {
        this.total_bug_num = total_bug_num;
    }

    public String getTotal_bug_percent() {
        return total_bug_percent;
    }

    public void setTotal_bug_percent(String total_bug_percent) {
        this.total_bug_percent = total_bug_percent;
    }
}
