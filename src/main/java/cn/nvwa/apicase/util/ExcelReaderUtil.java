package cn.nvwa.apicase.util;

import cn.nvwa.apicase.entity.ExcelInfo;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;


/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO 读取Excel内容到数据库
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class ExcelReaderUtil {
    private static Logger logger = Logger.getLogger(ExcelReaderUtil.class.getName()); // 日志打印类

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    private static String interface_name_value = "";
    private static String interface_url_value = "";
    private static String method_value = "";
    private static String common_ground_value = "";
    private static String bug_type_value = "";



    /**
     * 根据文件后缀名类型获取对应的工作簿对象
     * @param inputStream 读取文件的输入流
     * @param fileType 文件后缀名类型（xls或xlsx）
     * @return 包含文件数据的工作簿对象
     * @throws IOException
     */
    public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (fileType.equalsIgnoreCase(XLS)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase(XLSX)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    /**
     * 读取Excel文件内容
     * @param filePath 要读取的Excel文件所在路径
     * @return 读取结果列表，读取失败时返回null
     */
    public static List<ExcelInfo> readExcel(String filePath) {


        Workbook workbook = null;
        FileInputStream inputStream = null;

        String fileName = null;
        String sheetIndex = null;

        if(filePath.endsWith(".xlsx") || filePath.endsWith(".xls")){
            fileName = filePath;
        }else{
            String[] splitFilePath = filePath.split(" ");
            fileName = splitFilePath[0];
            sheetIndex = splitFilePath[1];
        }

        try {
            // 获取Excel后缀名
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
            // 获取Excel文件
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warn("指定的Excel文件不存在！");
                return null;
            }

            // 获取Excel工作簿
            inputStream = new FileInputStream(excelFile);
            workbook = getWorkbook(inputStream, fileType);

            // 读取excel中的数据
            List<ExcelInfo> resultDataList = parseExcel(workbook,sheetIndex);
            System.out.println("正在读取excel用例数据，请耐心等待～");

            return resultDataList;
        } catch (Exception e) {
            logger.warn("解析Excel失败，文件名：" + fileName + " 错误信息：" + e.getMessage());
            return null;
        } finally {
            try {
                if (null != workbook) {
                    workbook.close();
                }
                if (null != inputStream) {
                    inputStream.close();
                }
            } catch (Exception e) {
                logger.warn("关闭数据流出错！错误信息：" + e.getMessage());
                return null;
            }
        }
    }

    /**
     * 解析Excel数据
     * @param workbook Excel工作簿对象
     * @return 解析结果
     */
    private static List<ExcelInfo> parseExcel(Workbook workbook,String sheetIndex) {
        List<ExcelInfo> resultDataList = new ArrayList<>();

        if("".equals(sheetIndex) || null == sheetIndex){
            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = workbook.getSheetAt(sheetNum);

                // 校验sheet是否合法
                if (sheet == null) {
                    continue;
                }

                // 获取第一行数据
                int firstRowNum = sheet.getFirstRowNum();//从0开始
                Row firstRow = sheet.getRow(firstRowNum);//从0开始


                if (null == firstRow) {
                    logger.warn("解析Excel失败，在第一行没有读取到任何数据！");
                }

                // 解析每一行的数据，构造数据对象
                int rowStart = firstRowNum + 1;
                int rowEnd = sheet.getPhysicalNumberOfRows();


                for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                    Row row = sheet.getRow(rowNum);


                    if (null == row) {
                        continue;
                    }

                    ExcelInfo resultData = convertRowToData(row,firstRow);
                    if (null == resultData) {
                        logger.warn("第 " + row.getRowNum() + "行数据不合法，已忽略！");
                        continue;
                    }
                    resultDataList.add(resultData);
                }
            }
        }else{
            String[] sheetIndexStr = sheetIndex.split(",");
            for(String indexStr : sheetIndexStr){
                int index = Integer.parseInt(indexStr)-1;
                if( index < workbook.getNumberOfSheets()){
                    Sheet sheet = workbook.getSheetAt(index);

                    // 获取第一行数据
                    int firstRowNum = sheet.getFirstRowNum();//从0开始
                    Row firstRow = sheet.getRow(firstRowNum);//从0开始


                    if (null == firstRow) {
                        logger.warn("解析Excel失败，在第一行没有读取到任何数据！");
                    }

                    // 解析每一行的数据，构造数据对象
                    int rowStart = firstRowNum + 1;
                    int rowEnd = sheet.getPhysicalNumberOfRows();


                    for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                        Row row = sheet.getRow(rowNum);


                        if (null == row) {
                            continue;
                        }

                        ExcelInfo resultData = convertRowToData(row,firstRow);
                        if (null == resultData) {
                            logger.warn("第 " + row.getRowNum() + "行数据不合法，已忽略！");
                            continue;
                        }
                        resultDataList.add(resultData);
                    }
                }else if(index < 0 || index >= workbook.getNumberOfSheets()){
                    logger.warn("sheet下标不合法");
                }
            }
        }



        return resultDataList;
    }

    /**
     * 将单元格内容转换为字符串
     * @param cell
     * @return
     */
    public static String convertCellValueToString(Cell cell) {
        if(cell==null){
            return null;
        }
        String returnValue = null;


        switch (cell.getCellTypeEnum()) {
            case NUMERIC:   //数字NUMERIC
                Double doubleValue = cell.getNumericCellValue();

                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }

    /**
     * 提取每一行中需要的数据，构造成为一个结果数据对象
     * @param row 行数据
     * @return 解析后的行数据对象，行数据错误时返回null
     */
    public static ExcelInfo convertRowToData(Row row,Row firstRow) {
        ExcelInfo resultData = new ExcelInfo();

        Cell cell;
        int cellNum = 0;

        short firstRowLastCellNum = firstRow.getLastCellNum();

        for(int i=0;i<firstRowLastCellNum;i++){
            Cell firstRowCell = firstRow.getCell(i);
            if(firstRowCell.toString().equals("用例")){
                //获取用例
                cell = row.getCell(cellNum++);
                String test_case = convertCellValueToString(cell);
                if(null == test_case || "".equals(test_case)){
                    //用例为空
                    resultData.setTest_case(null);
                }else {
                    resultData.setTest_case(test_case);
                }
            }else if(firstRowCell.toString().equals("接口名称")){
                // 获取接口名称
                cell = row.getCell(cellNum++);
                String interface_name = convertCellValueToString(cell);
                if(null == interface_name || "".equals(interface_name)){
                    //接口名称为空
                    resultData.setInterface_name(interface_name_value);
                }else{
                    interface_name_value = interface_name;
                    resultData.setInterface_name(interface_name_value);
                }
            }else if(firstRowCell.toString().equals("接口路径")){
                //获取接口路径
                cell = row.getCell(cellNum++);
                String interface_url = convertCellValueToString(cell);
                if(null == interface_url || "".equals(interface_url)){
                    //接口路径为空
                    resultData.setInterface_url(interface_url_value); //2,3,4
                }else{
                    interface_url_value = interface_url;
                    resultData.setInterface_url(interface_url_value);
                }
            }else if(firstRowCell.toString().equals("方法")){
                //获取方法
                cell = row.getCell(cellNum++);
                String method = convertCellValueToString(cell);
                if(null == method || "".equals(method)){
                    //方法为空
                    resultData.setMethod(method_value);
                }else{
                    method_value = method;
                    resultData.setMethod(method_value);
                }
            }else if(firstRowCell.toString().equals("参数")){
                //获取参数
                cell = row.getCell(cellNum++);
                String parameter = convertCellValueToString(cell);
                if(null == parameter || "".equals(parameter)){
                    //参数为空
                    resultData.setParameter(null);
                }else{
                    resultData.setParameter(parameter);
                }
            }else if(firstRowCell.toString().equals("优先级")){
                //获取优先级
                cell = row.getCell(cellNum++);
                String priority = convertCellValueToString(cell);
                if(null == priority || "".equals(priority)){
                    //优先级为空
                    resultData.setPriority(null);
                }else{
                    resultData.setPriority(priority);
                }
            }else if(firstRowCell.toString().equals("共同点")){
                //获取共同点
                cell = row.getCell(cellNum++);
                String common_ground = convertCellValueToString(cell);
                if(null == common_ground || "".equals(common_ground)){
                    //共同点为空
                    resultData.setCommon_ground(common_ground_value);
                }else{
                    common_ground_value = common_ground;
                    resultData.setCommon_ground(common_ground_value);
                }
            }else if(firstRowCell.toString().equals("测试重点")){
                //获取测试重点
                cell = row.getCell(cellNum++);
                String test_emphasis = convertCellValueToString(cell);
                resultData.setTest_emphasis(test_emphasis);
            }else if(firstRowCell.toString().equals("预期结果")){
                //获取预期结果
                cell = row.getCell(cellNum++);
                String result_expected = convertCellValueToString(cell);
                if(null == result_expected || "".equals(result_expected)){
                    //预期结果为空
                    resultData.setResult_expected(null);
                }else{
                    resultData.setResult_expected(result_expected);
                }
            }else if(firstRowCell.toString().equals("实际结果")){
                //获取实际结果
                cell = row.getCell(cellNum++);
                String result_actual = convertCellValueToString(cell);
                if(null == result_actual || "".equals(result_actual)){
                    //实际结果为空
                    resultData.setResult_actual(null);
                }else{
                    resultData.setResult_actual(result_actual);
                }
            }else if(firstRowCell.toString().equals("问题类型")){
                // 问题类型
                cell = row.getCell(cellNum++);
                String bug_type = convertCellValueToString(cell);
                if (null == bug_type || "".equals(bug_type)) {
                    // 问题类型为空
                    resultData.setBug_type("测试通过");
                } else {
                    bug_type_value = bug_type;
                    resultData.setBug_type(bug_type_value);
                }
            }else if(firstRowCell.toString().equals("修改情况")){
                //修改情况
                cell = row.getCell(cellNum++);
                String alter_situation = convertCellValueToString(cell);
                if(null == alter_situation || "".equals(alter_situation)){
                    //修改情况为空
                    resultData.setAlter_situation(null);
                }else{
                    resultData.setAlter_situation(alter_situation);
                }
            }else {
                cellNum++;
                continue;
            }

        }

        return resultData;
    }

}
