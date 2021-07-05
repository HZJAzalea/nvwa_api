package cn.nvwa.apicase.main;

import cn.nvwa.apicase.dao.ExcelDao;

import cn.nvwa.apicase.entity.*;
import cn.nvwa.apicase.util.ExcelReaderUtil;
import cn.nvwa.apicase.util.ExcelWriterUtil;

import java.util.List;

/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class TestMain {
    public static void main(String[] args) {



        //写入mysql的文件路径
//      String excelFilePathAndIndex = "/Users/huangzhijuan/Documents/测试工作/接口测试/1.中台/限流服务/配置服务/配置服务.xlsx";

        String excelFilePathAndIndex = null;
        if(args.length == 1){
            //写入mysql的文件路径（不包含sheetIndex）
            excelFilePathAndIndex = args[0];
        }else{
            //写入mysql的文件路径（包含sheetIndex）
            excelFilePathAndIndex = args[0] + " " + args[1];
        }


        //写入mysql的文件路径
        String excelFilePath = excelFilePathAndIndex.split(" ")[0];


        //写入excel的文件路径
        String exportFilePath = null;
        if(excelFilePath.endsWith(".xlsx")){
            exportFilePath = excelFilePath.replace(".xlsx","-bug率统计.xlsx");
        }else if(excelFilePath.endsWith(".xls")){
            exportFilePath = excelFilePath.replace(".xls","-bug率统计.xls");
        }

        //excel数据入mysql
        //读取Excel文件内容
        List<ExcelInfo> excelInfoList = ExcelReaderUtil.readExcel(excelFilePathAndIndex);
        ExcelDao excelDao = new ExcelDao();
        excelDao.saveExcelInfo(excelInfoList);


        //mysql入excel
        List<MysqlToExcel> mysqlToExcelList = ExcelDao.dbcpQuery();
        List<MysqlToExcelTotal> mysqlToExcelTotalList = ExcelDao.dbcpQueryTotal();
        ExcelWriterUtil.excelWriter(exportFilePath,mysqlToExcelList,mysqlToExcelTotalList);

    }
}
