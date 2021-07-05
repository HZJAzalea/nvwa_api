package cn.nvwa.apicase.dao;

import cn.nvwa.apicase.entity.ExcelInfo;
import cn.nvwa.apicase.entity.MysqlToExcel;
import cn.nvwa.apicase.entity.MysqlToExcelTotal;
import cn.nvwa.apicase.util.DbcpUtil;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.util.*;

/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class ExcelDao {
    private static PreparedStatement ps = null;
    static Set<String> urlSet = new HashSet<>();
    private static int totalTestCaseNum = 0;
    private static int totalBugNum = 0;


    /**
     * @Title: delData
     * @Description: 删除表中已存在的数据
     * @param: @param interface_url
     * @param: @return
     * @return: int[]
     * @throws
     */
    public int[] delData(String interface_url){
        PreparedStatement psDel = null;

        String sql = "delete from interface_testcase where interface_url = " + "\"" + interface_url + "\"";
        int[] count = null;

        Connection connection = null;
        try{
            connection = DbcpUtil.getConnection();
            psDel = connection.prepareStatement(sql);


            psDel.addBatch();
            //批量删除
            count = psDel.executeBatch();

        }catch(Exception e){
            e.printStackTrace();
        }finally {
            DbcpUtil.close(null,psDel,connection);
        }

        return count;
    }

    /**
     * @Title: isExist
     * @Description: 判断表中是否存在该数据
     * @param: @param interface_url
     * @param: @return
     * @return: boolean
     * @throws
     */
    public static boolean isExist(String interface_url){
        String sql_findByUrl = "select * from interface_testcase where interface_url = " + "\""+interface_url + "\"";
        Connection connection = null;
        PreparedStatement psExists = null;
        ResultSet rs = null;
        boolean exist = false;

        try {
            connection = DbcpUtil.getConnection();
            psExists = connection.prepareStatement(sql_findByUrl);
            rs = psExists.executeQuery();
            if(rs.next() == true){
                exist = true;
            }else{
                exist = false;
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }finally {
            DbcpUtil.close(rs,psExists,connection);
        }

        return exist;
    }


    /**
     * @Title: saveExcelInfo
     * @Description: 基于jdbc保存用例数据入库
     * @param: @param excelInfoList
     * @param: @return
     * @return: int[]
     * @throws
     */
    public int[] saveExcelInfo(List<ExcelInfo> excelInfoList){
        String sql = "insert into interface_testcase(test_case,interface_name,interface_url,method,parameter,priority,common_ground,test_emphasis,result_expected,result_actual,bug_type,alter_situation) values (?,?,?,?,?,?,?,?,?,?,?,?)";
        int[] count = null;

        Connection connection = null;
        try{
            connection = DbcpUtil.getConnection();
            ps = connection.prepareStatement(sql);


            if(excelInfoList != null && !excelInfoList.isEmpty()){
                for(ExcelInfo excelInfo :excelInfoList){

                    String interface_url = excelInfo.getInterface_url();

                    urlSet.add(interface_url);

                    //insert之前先判断库中是否存在，如果存在则删除原数据
                    for(String url : urlSet){
                        if(isExist(url) == true){
                            delData(url);
                        }
                    }
                    ps.setString(1, excelInfo.getTest_case());
                    ps.setString(2, excelInfo.getInterface_name());
                    ps.setString(3, interface_url);
                    ps.setString(4, excelInfo.getMethod());
                    ps.setString(5, excelInfo.getParameter());
                    ps.setString(6, excelInfo.getPriority());
                    ps.setString(7, excelInfo.getCommon_ground());
                    ps.setString(8, excelInfo.getTest_emphasis());
                    ps.setString(9, excelInfo.getResult_expected());
                    ps.setString(10, excelInfo.getResult_actual());
                    ps.setString(11, excelInfo.getBug_type());
                    ps.setString(12, excelInfo.getAlter_situation());
                    ps.addBatch();
                }
            }

            //批量插入
            count = ps.executeBatch();

            //插入之后重置排序
            resetSort();


        }catch(Exception e){
            e.printStackTrace();
        }finally {
            DbcpUtil.close(null,ps,connection);
        }

        return count;
    }

    /**
     * @Title: resetSort
     * @Description: 重置排序
     * @param: @param
     * @param: @return
     * @return:
     * @throws
     */
    public static void resetSort(){

        Connection connection = null;

        try{
            connection = DbcpUtil.getConnection();
            connection.setAutoCommit(false);
            ps = connection.prepareStatement("SET @i=0");
            ps.executeUpdate();
            ps = connection.prepareStatement("UPDATE interface_testcase SET id=(@i:=@i+1)");
            ps.executeUpdate();
            ps = connection.prepareStatement("ALTER TABLE interface_testcase AUTO_INCREMENT=0");
            ps.executeUpdate();
            connection.commit();


        }catch (Exception e){
            e.printStackTrace();
        }finally {
            DbcpUtil.close(null,ps,connection);
        }
    }
    /**
     * @Title: dbcpQuery
     * @Description: 基于jdbc进行查询，查询出接口名称、接口路径、接口用例数、bug类型、bug类型统计数（明细表）
     * @param: @return
     * @return: List<MysqlToExcel>
     * @throws
     */
    public static List<MysqlToExcel> dbcpQuery(){
        Map<String, Integer> testCaseNum = dbcpQueryTestCaseNum();
        Map<String,Integer> totalBugNum = dbcpQueryPerApiTotalBugNum();
        //初始化
        ResultSet rs = null;
        Connection connection = null;
        List<MysqlToExcel> list = new ArrayList<>();

        try {
            connection = DbcpUtil.getConnection();
            for(String url : urlSet){

                    String querySql = "select interface_name,interface_url,bug_type,count(bug_type) bug_type_num from interface_testcase where bug_type not in(\"测试通过\") and interface_url=" + "\"" + url + "\"" + " group by interface_name,interface_url,bug_type";

                    ps = connection.prepareStatement(querySql);
                    //执行select sql
                    rs = ps.executeQuery();


                    //ResultSet结果集处理
                    while (rs.next()){
                        MysqlToExcel mysqlToExcel = new MysqlToExcel();
                        //接口名称
                        String interface_name = rs.getString("interface_name");
                        mysqlToExcel.setInterface_name(interface_name);
                        //接口路径
                        String interface_url = rs.getString("interface_url");
                        mysqlToExcel.setInterface_url(interface_url);
                        //每个接口的总用例数（明细表）
                        int testcase_num = testCaseNum.get(interface_url);
                        mysqlToExcel.setTestcase_num(testcase_num);
                        //每个接口的bug总数（明细表）
                        int total_bug_num = totalBugNum.get(interface_url);
                        mysqlToExcel.setTotal_bug_num(total_bug_num);
                        //bug总占比（明细表）
                        if(testcase_num != 0){
                            double totalPercent = (double) total_bug_num/testcase_num * 100;
                            DecimalFormat fnum = new DecimalFormat("##0.00");
                            String totalPercentStr = fnum.format(totalPercent);
                            mysqlToExcel.setTotal_bug_percent(totalPercentStr + "%");
                        }
                        //bug类型（明细表）
                        String bug_type = rs.getString("bug_type");
                        mysqlToExcel.setBug_type(bug_type);
                        //bug类型统计数目（明细表）
                        int bug_type_num = rs.getInt("bug_type_num");
                        mysqlToExcel.setBug_type_num(bug_type_num);
                        //bug类型百分比（明细表）
                        if(testcase_num != 0){
                            double percent =(double) bug_type_num/testcase_num * 100;
                            DecimalFormat fnum = new DecimalFormat("##0.00");
                            String percentStr =fnum.format(percent);
                            mysqlToExcel.setPercent(percentStr + "%");
                        }

                        list.add(mysqlToExcel);
                    }

            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally{
            DbcpUtil.close(rs,ps,connection);
        }
        return list;
    }

    /**
     * @Title: dbcpQuery
     * @Description: 基于jdbc进行查询，获取各个接口用例数目（write明细表）
     * @param: @return
     * @return: Map<String,Integer>
     * @throws
     */
    public static Map<String,Integer> dbcpQueryTestCaseNum(){
        //初始化
        ResultSet rs = null;
        Connection connection = null;
        Map<String,Integer> map = new HashMap<>();

        try {
            connection = DbcpUtil.getConnection();
            for(String url : urlSet){

                    String sql = "select interface_url,count(*) interface_url_num from interface_testcase where test_emphasis not in(\"\") and interface_url=" + "\"" + url + "\"" + " group by interface_url";

                    ps = connection.prepareStatement(sql);
                    //执行select sql
                    rs = ps.executeQuery();


                    while (rs.next()){
                        String interface_url = rs.getString("interface_url");

                        int interface_url_num = rs.getInt("interface_url_num");

                        map.put(interface_url,interface_url_num);

                        //得到接口总用例数（汇总表）
                        totalTestCaseNum += interface_url_num;

                    }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally{
            DbcpUtil.close(rs,ps,connection);
        }
        return map;
    }

    /**
     * @Title: dbcpQuery
     * @Description: 基于jdbc进行查询，获取各个接口用例总bug数目（明细用）
     * @param: @return
     * @return: Map<String,Integer>
     * @throws
     */
    public static Map<String,Integer> dbcpQueryPerApiTotalBugNum(){
        //初始化
        ResultSet rs = null;
        Connection connection = null;
        Map<String,Integer> map = new HashMap<>();

        try {
            connection = DbcpUtil.getConnection();
            for(String url : urlSet){

                    String perApiSql = "select interface_url,count(*) per_api_total_bug_num from interface_testcase where bug_type not in(\"测试通过\") and interface_url=" + "\"" + url + "\"" + " group by interface_url";

                    ps = connection.prepareStatement(perApiSql);
                    //执行select sql
                    rs = ps.executeQuery();


                    while (rs.next()){
                        String interface_url = rs.getString("interface_url");
                        int per_api_total_bug_num = rs.getInt("per_api_total_bug_num");

                        map.put(interface_url,per_api_total_bug_num);
                        //得到接口总bug数（汇总表）
                        totalBugNum += per_api_total_bug_num;

                    }

            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally{
            DbcpUtil.close(rs,ps,connection);
        }
        return map;
    }

    /**
     * @Title: dbcpQueryTotal
     * @Description: 基于jdbc进行查询，获取总bug类型及其数目（汇总用）
     * @param: @return
     * @return: List<MysqlToExcelTotal>
     * @throws
     */
    public static List<MysqlToExcelTotal> dbcpQueryTotal() {
        //初始化
        ResultSet rs = null;
        Connection connection = null;
        List<MysqlToExcelTotal> list = new ArrayList<>();
        Map<String,Integer> bugTypeAndTotalBugNumMap = new HashMap<>();

        try {
            connection = DbcpUtil.getConnection();

            for(String url:urlSet){

                String sql = "select bug_type,count(bug_type) total_bug_type_num from interface_testcase where bug_type not in(\"测试通过\") and interface_url="+ "\"" + url + "\"" +" group by bug_type";

                ps = connection.prepareStatement(sql);
                rs = ps.executeQuery();
                while (rs.next()) {

                    int total_bug_type_num_map = 0;
                    //问题类型
                    String bug_type_map = rs.getString("bug_type");
                    int total_bug_type_num_temp = rs.getInt("total_bug_type_num");
                    if(bugTypeAndTotalBugNumMap.containsKey(bug_type_map)){
                        total_bug_type_num_map = bugTypeAndTotalBugNumMap.get(bug_type_map) + total_bug_type_num_temp;
                        bugTypeAndTotalBugNumMap.put(bug_type_map,total_bug_type_num_map);
                    }else {
                        total_bug_type_num_map = total_bug_type_num_temp;
                        bugTypeAndTotalBugNumMap.put(bug_type_map,total_bug_type_num_map);
                    }

                }

            }

            for(Map.Entry<String,Integer> m : bugTypeAndTotalBugNumMap.entrySet()){
                MysqlToExcelTotal mysqlToExcelTotal = new MysqlToExcelTotal();
                //总接口用例数
                mysqlToExcelTotal.setTotal_tesecase_num(totalTestCaseNum);
                //bug总占比
                if(totalTestCaseNum != 0){
                    double percent = (double) totalBugNum/totalTestCaseNum * 100;
                    DecimalFormat fnum = new DecimalFormat("##0.00");
                    String totalPercentStr = fnum.format(percent);
                    mysqlToExcelTotal.setBug_total_percent(totalPercentStr + "%");
                }
                String bug_type = m.getKey();
                mysqlToExcelTotal.setTotal_bug_type(bug_type);
                int total_bug_type_num = m.getValue();
                mysqlToExcelTotal.setTotal_bug_type_bum(total_bug_type_num);
                //各bug类型总占比
                if(totalTestCaseNum != 0){
                    double percent =(double) total_bug_type_num/totalTestCaseNum * 100;
                    DecimalFormat fnum = new DecimalFormat("##0.00");
                    String percentStr =fnum.format(percent);
                    mysqlToExcelTotal.setTotal_percent(percentStr + "%");

                }

                list.add(mysqlToExcelTotal);

            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            DbcpUtil.close(rs, ps, connection);
        }
        return list;
    }

}
