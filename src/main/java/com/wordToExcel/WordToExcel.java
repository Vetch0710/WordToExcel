package com.wordToExcel;

import com.wordToExcel.entity.TableData;
import com.wordToExcel.tool.ExcelTools;
import com.wordToExcel.tool.WordTools;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Properties;

/**
 * @author Vetch
 * @Description
 * @create 2021-03-22-14:12
 */
public class WordToExcel {
    public static void main(String[] args) throws Exception {
//        "C:\\Users\\75440\\Desktop\\统计信息管理平台过程统计技术指标卡片(1).docx"
//        "C:\\Users\\75440\\Desktop\\统计信息管理平台指标体系业务指标卡片 (1).docx"
        WordToExcel.wordToExcel("C:\\Users\\75440\\Desktop\\统计信息管理平台指标体系业务指标卡片 (1).docx",
                "C:\\Users\\75440\\Desktop\\指标体系梳理表 (1) (1).xlsx",
                "src/main/resources/ExcelRule.properties",
                "src/main/resources/WordRule.properties",
                1, 3
        );
    }


    /**
     * Description: word表格数据导入Excel
     *
     * @param wordUrl
     * @param excelUrl
     * @param excelRuleUrl
     * @param wordRuleUrl
     * @param sheetNum  操作的sheet页
     * @param startRow  开始行
     * @Return: void
     * @Date: 2021/03/22 15:31
     */
    public static void wordToExcel(String wordUrl, String excelUrl, String excelRuleUrl, String wordRuleUrl, int sheetNum, int startRow) throws Exception {
        long startTime = System.currentTimeMillis();
        //读取配置文件
        InputStream input = new BufferedInputStream(new FileInputStream(wordRuleUrl));
        Properties wordProperties = new Properties();
        wordProperties.load(input);
        InputStream input2 = new BufferedInputStream(new FileInputStream(excelRuleUrl));
        Properties excelProperties = new Properties();
        excelProperties.load(input2);
        //读取word表格数据
        List<TableData> tableData = WordTools.tableInWord(wordUrl, wordProperties);

//        tableData.forEach(System.out::println);
        ExcelTools excelTool = new ExcelTools();
        //读取表头
        List<Object> excelValues = excelTool.getExcelTitles(excelUrl, sheetNum, excelProperties);
//        获取Excel数据行数
//        int row = excelTool.getRow(excelUrl, sheetNum);
        //导入数据
        ExcelTools.appendDateToExcel(excelUrl, tableData, startRow - 1, sheetNum, excelValues);
        long endTime = System.currentTimeMillis();

        System.out.println("程序运行时间：" + (endTime - startTime) + "ms");
    }
}
