package cn.nvwa.apicase.util;

import cn.nvwa.apicase.entity.MysqlToExcel;
import cn.nvwa.apicase.entity.MysqlToExcelTotal;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


/**
 * @ClassName: 1
 * @Author: huangzhijuan
 * @Description: TODO 写数据到excel
 * @Date: Create in 1 1
 * @Version 1.0
 */
public class ExcelWriterUtil {
    private static Logger logger = Logger.getLogger(ExcelWriterUtil.class.getName());

    private static List<String> CELL_HEADS; //明细列头
    private static List<String> CELL_HEADS_SUM;//汇总表头
    // 类装载时就载入指定好的列头信息，如有需要，可以考虑做成动态生成的列头
    static {
        CELL_HEADS = new ArrayList<>();
        CELL_HEADS.add("接口名称");
        CELL_HEADS.add("接口路径");
        CELL_HEADS.add("总用例数");
        CELL_HEADS.add("bug总数");
        CELL_HEADS.add("bug总占比");
        CELL_HEADS.add("bug类型");
        CELL_HEADS.add("bug类型统计数");
        CELL_HEADS.add("bug类型百分比");

        CELL_HEADS_SUM = new ArrayList<>();
        CELL_HEADS_SUM.add("总用例数");
        CELL_HEADS_SUM.add("bug总占比");
        CELL_HEADS_SUM.add("bug类型");
        CELL_HEADS_SUM.add("bug类型统计总数");
        CELL_HEADS_SUM.add("各bug类型占比");
    }

    /**
     * 明细sheet表合并单元格
     * @param mysqlToExcels 数据列表
     * @param sheet sheet对象
     * @return
     */

    public static void mergeRow(List<MysqlToExcel> mysqlToExcels, Sheet sheet){

        int size = mysqlToExcels.size();

        int j=0;
        for(int i=0;i<size;i++){

            MysqlToExcel mysqlToExcel=mysqlToExcels.get(i);
            String url=mysqlToExcel.getInterface_url();
            String urlNext="";
            if (size>(i+1)){
                MysqlToExcel mysqlToExcelNext=mysqlToExcels.get(i+1);
                urlNext=mysqlToExcelNext.getInterface_url();
            }

            if (url.equals(urlNext)){
                j=j+1;
                continue;
            }

            int first=i+1-j;
            int last=i+1;

            j=0;
            if (first==last){
                continue;
            }
            CellRangeAddress callRangeAddress = new CellRangeAddress(first,last,0,0);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress);

            CellRangeAddress callRangeAddress1 = new CellRangeAddress(first,last,1,1);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress1);

            CellRangeAddress callRangeAddress2 = new CellRangeAddress(first,last,2,2);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress2);

            CellRangeAddress callRangeAddress3 = new CellRangeAddress(first,last,3,3);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress3);

            CellRangeAddress callRangeAddress4 = new CellRangeAddress(first,last,4,4);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress4);

        }

    }

    /**
     * 汇总sheet表合并单元格
     * @param mysqlToExcelTotalList 数据列表
     * @param sheet sheet对象
     * @return
     */
    public static void mergeRowTotal(List<MysqlToExcelTotal> mysqlToExcelTotalList, Sheet sheet){

        int size = mysqlToExcelTotalList.size();

        int j=0;
        for(int i=0;i<size;i++){

            MysqlToExcelTotal mysqlToExcelTotal = mysqlToExcelTotalList.get(i);
            int total_tesecase_num = mysqlToExcelTotal.getTotal_tesecase_num();
            int totalTestcaseNumNext = 0;
            if (size>(i+1)){
                MysqlToExcelTotal mysqlToExcelTotalNext = mysqlToExcelTotalList.get(i + 1);
                totalTestcaseNumNext=mysqlToExcelTotalNext.getTotal_tesecase_num();
            }

            if (total_tesecase_num == totalTestcaseNumNext){
                j=j+1;
                continue;
            }

            int first=i+1-j;
            int last=i+1;

            j=0;
            if (first==last){
                continue;
            }
            CellRangeAddress callRangeAddress = new CellRangeAddress(first,last,0,0);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress);

            CellRangeAddress callRangeAddress1 = new CellRangeAddress(first,last,1,1);//起始行,结束行,起始列,结束列
            sheet.addMergedRegion(callRangeAddress1);

        }

    }

    /**
     * 生成Excel并写入数据信息
     * @param mysqlToExcels 数据列表
     * @return 写入数据后的工作簿对象
     */
    public static Workbook exportData(List<MysqlToExcel> mysqlToExcels,List<MysqlToExcelTotal> mysqlToExcelTotalList){
        // 生成xlsx的Excel
        Workbook workbook = new SXSSFWorkbook();

        //单元格左对齐
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置自动换行
        cellStyle.setWrapText(true);
        //加边框
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());//下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框

        // 生成明细Sheet表，写入第一行的列头
        Sheet sheet = buildDataSheet(workbook);
        mergeRow(mysqlToExcels,sheet);

        System.out.println("正在生成明细表，请耐心等待～");

        //构建每行的数据内容
        int rowNum = 1;
        for(Iterator<MysqlToExcel> it = mysqlToExcels.iterator(); it.hasNext();){
            MysqlToExcel data = it.next();
            if(data == null){
                continue;
            }
            //输出行数据
            Row row = sheet.createRow(rowNum++);

            int cellNum = convertDataToRow(data, row);
            for(int i=0;i<cellNum;i++){
                Cell cell = row.getCell(i);
                cell.setCellStyle(cellStyle);
            }

        }

        // 生成汇总Sheet表，写入第一行的列头
        Sheet sheetTotal = buildDataSheetTotal(workbook);
        mergeRowTotal(mysqlToExcelTotalList,sheetTotal);
        System.out.println("正在生成汇总表，请耐心等待～");
        //构建每行的数据内容
        int rowNumTotal = 1;
        for(Iterator<MysqlToExcelTotal> it = mysqlToExcelTotalList.iterator(); it.hasNext();){
            MysqlToExcelTotal data = it.next();
            if(data == null){
                continue;
            }
            //输出行数据
            Row row = sheetTotal.createRow(rowNumTotal++);
            int cellNum = convertDataToRowTotal(data, row);
            for(int i=0;i<cellNum;i++){
                Cell cell = row.getCell(i);
                cell.setCellStyle(cellStyle);
            }

        }
        return workbook;
    }

    /**
     * 生成明细sheet表，并写入第一行数据（列头）
     * @param workbook 工作簿对象
     * @return 已经写入列头的Sheet
     */
    private static Sheet buildDataSheet(Workbook workbook){
        Sheet sheet = workbook.createSheet();
        workbook.setSheetName(0,"明细");
        // 设置列头宽度
        for(int i=0;i<CELL_HEADS.size();i++){
            sheet.setColumnWidth(i,4000);
        }
        // 设置默认行高
        sheet.setDefaultRowHeight((short)400);
        // 构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(workbook);
        // 写入第一行各列的数据
        Row head = sheet.createRow(0);
        for(int i=0;i<CELL_HEADS.size();i++){
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * 生成汇总sheet表，并写入第一行数据（列头）
     * @param workbook 工作簿对象
     * @return 已经写入列头的Sheet
     */
    private static Sheet buildDataSheetTotal(Workbook workbook){
        Sheet sheet = workbook.createSheet();
        workbook.setSheetName(1,"汇总");

        // 设置列头宽度
        for(int i=0;i<CELL_HEADS_SUM.size();i++){
            sheet.setColumnWidth(i,4500);
        }
        // 设置默认行高
        sheet.setDefaultRowHeight((short)400);
        // 构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(workbook);
        // 写入第一行各列的数据
        Row head = sheet.createRow(0);
        for(int i=0;i<CELL_HEADS_SUM.size();i++){
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS_SUM.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * 设置第一行列头的样式
     * @param workbook 工作簿对象
     * @return 单元格样式对象
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 明细sheet 将数据转换成行
     * @param mysqlToExcel 源数据
     * @param row 行对象
     * @return
     */
    private static int convertDataToRow(MysqlToExcel mysqlToExcel, Row row){


        int cellNum = 0;
        Cell cell;


        // 接口名称
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == mysqlToExcel.getInterface_name() ? "" : mysqlToExcel.getInterface_name());
        // 接口路径
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == mysqlToExcel.getInterface_url() ? "" : mysqlToExcel.getInterface_url());



        //总用例数
        cell = row.createCell(cellNum++);
        //cell.setCellValue(testcase_num);
        cell.setCellValue(mysqlToExcel.getTestcase_num());

        //bug总数
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcel.getTotal_bug_num());

        //bug总占比
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcel.getTotal_bug_percent());

        //bug类型
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == mysqlToExcel.getBug_type() ? "" : mysqlToExcel.getBug_type());


        //bug类型统计数
        cell = row.createCell(cellNum++);
        //cell.setCellValue(bug_type_num);
        cell.setCellValue(mysqlToExcel.getBug_type_num());

        //bug类型百分比
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcel.getPercent());

        return cellNum;


    }

    /**
     * 汇总sheet 将数据转换成行
     * @param mysqlToExcelTotal 源数据
     * @param row 行对象
     * @return
     */
    private static int convertDataToRowTotal(MysqlToExcelTotal mysqlToExcelTotal, Row row){


        int cellNum = 0;
        Cell cell;

        // 总用例数
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcelTotal.getTotal_tesecase_num());

        //bug总占比
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcelTotal.getBug_total_percent());

        //问题类型
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == mysqlToExcelTotal.getTotal_bug_type() ? "" : mysqlToExcelTotal.getTotal_bug_type());


        //bug类型统计总数
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcelTotal.getTotal_bug_type_bum());

        //问题类型百分比
        cell = row.createCell(cellNum++);
        cell.setCellValue(mysqlToExcelTotal.getTotal_percent());

        return cellNum;

    }


    /**
     * 以文件形式输出excel
     * @param
     * @param mysqlToExcelList 数据列表集合
     * @return
     */
    public static void excelWriter(String exportFilePath,List<MysqlToExcel> mysqlToExcelList,List<MysqlToExcelTotal> mysqlToExcelTotalList){


        // 写入数据到工作簿对象内
        Workbook workbook = ExcelWriterUtil.exportData(mysqlToExcelList,mysqlToExcelTotalList);

        // 以文件的形式输出工作簿对象
        FileOutputStream fileOut = null;
        try {
            File exportFile = new File(exportFilePath);
            if (!exportFile.exists()) {
                exportFile.createNewFile();
            }

            fileOut = new FileOutputStream(exportFilePath);
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            logger.warn("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                if (null != workbook) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.warn("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }

    }


}

