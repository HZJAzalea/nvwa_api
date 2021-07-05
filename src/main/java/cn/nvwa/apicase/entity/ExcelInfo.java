package cn.nvwa.apicase.entity;

/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class ExcelInfo {
    private static final long serialVersionUID=1L;

    private int id;
    private String test_case;
    private String interface_name;
    private String interface_url;
    private String method;
    private String parameter;
    private String priority;
    private String common_ground;
    private String test_emphasis;
    private String result_expected;
    private String result_actual;
    private String bug_type;
    private String alter_situation;



    public ExcelInfo(){}

    public ExcelInfo(String test_case, String interface_name, String interface_url, String method, String parameter, String priority, String common_ground, String test_emphasis, String result_expected, String result_actual, String bug_type, String alter_situation) {
        this.test_case = test_case;
        this.interface_name = interface_name;
        this.interface_url = interface_url;
        this.method = method;
        this.parameter = parameter;
        this.priority = priority;
        this.common_ground = common_ground;
        this.test_emphasis = test_emphasis;
        this.result_expected = result_expected;
        this.result_actual = result_actual;
        this.bug_type = bug_type;
        this.alter_situation = alter_situation;

    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getTest_case() {
        return test_case;
    }

    public void setTest_case(String test_case) {
        this.test_case = test_case;
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

    public String getMethod() {
        return method;
    }

    public void setMethod(String method) {
        this.method = method;
    }

    public String getParameter() {
        return parameter;
    }

    public void setParameter(String parameter) {
        this.parameter = parameter;
    }

    public String getPriority() {
        return priority;
    }

    public void setPriority(String priority) {
        this.priority = priority;
    }

    public String getCommon_ground() {
        return common_ground;
    }

    public void setCommon_ground(String common_ground) {
        this.common_ground = common_ground;
    }

    public String getTest_emphasis() {
        return test_emphasis;
    }

    public void setTest_emphasis(String test_emphasis) {
        this.test_emphasis = test_emphasis;
    }

    public String getResult_expected() {
        return result_expected;
    }

    public void setResult_expected(String result_expected) {
        this.result_expected = result_expected;
    }

    public String getResult_actual() {
        return result_actual;
    }

    public void setResult_actual(String result_actual) {
        this.result_actual = result_actual;
    }

    public String getBug_type() {
        return bug_type;
    }

    public void setBug_type(String bug_type) {
        this.bug_type = bug_type;
    }

    public String getAlter_situation() {
        return alter_situation;
    }



    public void setAlter_situation(String alter_situation) {
        this.alter_situation = alter_situation;
    }
}
