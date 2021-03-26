package com.wordToExcel.tool;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * @author Vetch
 * @Description
 * @create 2021-03-21-12:20
 */
public class ExcelTools {

    /**
     * Description:  读取表头
     *
     * @param fileUrl  Excel文件路径
     * @param sheetNum sheet数
     * @Return: java.util.List<java.lang.Object> 表头数据 有次级表头使用Map<String, Object>存储  没有则使用String
     * @Date: 2021/03/22 13:29
     */
    public static List<Object> getExcelTitles(String fileUrl, int sheetNum,Properties properties) throws Exception {
        List<Object> values = new ArrayList<>();
        //读取文件
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);

        Sheet sheet = workbook.getSheetAt(sheetNum - 1); //读取第几个工作表sheet
//        int rowNum = sheet.getPhysicalNumberOfRows();//有多少行
//        System.out.println(rowNum+"------------");
        Row row = sheet.getRow(0);//第i行
        if (row == null) {//过滤空行
            throw new Exception("表格无数据");
        }
        int colCount = sheet.getRow(0).getLastCellNum();//用表头去算有多少列，不然从下面的行计算列的话，空的就不算了
        int index = 0;
        while (index < colCount) {
            Cell cell = row.getCell(index);
            String cellValue;
            int isMerge = 0;
            //获取合并单元格列数
            if (cell != null) {
                isMerge = isMergedRegion(sheet, 0, cell.getColumnIndex());
            }
            //判断是否具有合并单元格
            if (isMerge != 0) {
                cellValue = getStringCellValue(cell);
                //获取表头数据对应对象中的属性名称
                String property = properties.getProperty(cellValue);
                //用来存储含有次级表头数据的表头  kay：一级表头 value：次级表头集合
                Map<String, Object> map = new LinkedHashMap<>();
                //获取次级表头集合
                List<Object> mergedValue = getMergedRegionObject(sheet, 1, cell.getColumnIndex(), isMerge,properties);
                map.put(property, mergedValue);
                //将该表头数据加入到List中
                values.add(map);
                index += isMerge + 1;
            } else { //无次级表头 直接添加
                cellValue = getStringCellValue(cell);
                String property = properties.getProperty(cellValue);
                values.add(property);
                index++;
            }
        }
        return values;
    }

    /**
     * Description: 获取次级表头数据
     *
     * @param sheet
     * @param nowRow      从哪行开始读取
     * @param nowColumn   从哪列开始读取
     * @param mergeColumn 合并的列数
     * @Return: java.util.List<java.lang.Object>    次级表头数据
     * @Date: 2021/03/22 13:45
     */
    public static List<Object> getMergedRegionObject(Sheet sheet, int nowRow, int nowColumn, int mergeColumn,Properties properties) {
        List<Object> result = new ArrayList<>();
        Row row = sheet.getRow(nowRow);//第i行
        //记录当前列数
        int index = nowColumn;
        while (index <= (nowColumn + mergeColumn)) {    //遍历列，遍历次数：当前列+合并列数
            Cell cell = row.getCell(index);
            //判断当前单元格是否还是合并列
            int mergedNum = isMergedRegion(sheet, nowRow, cell.getColumnIndex());
            if (mergedNum != 0) {  //是合并则使用递归 获取该合并单元格的次级表头数据
                Map<String, Object> map = new LinkedHashMap<>();
                List<Object> mergedRegionObject = getMergedRegionObject(sheet, nowRow + 1, cell.getColumnIndex(), mergedNum,properties);
                String mergedRegionValue = getStringCellValue(cell);
                map.put(properties.getProperty(mergedRegionValue), mergedRegionObject);
                result.add(map);
                index += mergedNum + 1;
            } else {    //不是合并列则，直接添加到List
                result.add(properties.getProperty(getStringCellValue(cell)));
                index++;
            }
        }
        return result;
    }


    /**
     * Description:  判断单元格是否合并
     *
     * @param sheet
     * @param row    行下标
     * @param column 列小标
     * @Return: int   合并的列数
     * @Date: 2021/03/22 13:26
     */
    public static int isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();

            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return lastColumn - firstColumn;
                }
            }
        }
        return 0;
    }


    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    public static String getStringCellValue(Cell cell) {
        String strCell = "";
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue()).trim();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue()).trim();
                break;
            case Cell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        return strCell;
    }


    /**
     * Description: 导入Excel文件
     *
     * @param excelUrl Excel文件路径
     * @param dateList 导入Excel文件 的数据
     * @param startRow 从哪行开始导入
     * @param SheetNum 读哪个sheet (输入从1开始)
     * @param titles   表头数据
     * @Return: boolean  导入Excel是否成功
     * @Date: 2021/03/22 13:20
     */
    public static boolean appendDateToExcel(String excelUrl, List<?> dateList,
                                            int startRow, int SheetNum, List<?> titles) {

        FileInputStream fs = null;
        FileOutputStream out = null;
        try {
            fs = new FileInputStream(excelUrl); // 获取Excel文件
            Workbook wb = WorkbookFactory.create(fs);
            fs.close();
            //得到Excel工作表对象  Excel中下标是从0开始
            Sheet sheet = wb.getSheetAt(SheetNum - 1);
            Row row;
            //获取TableData类
            Class<?> tableDataClass = dateList.get(0).getClass();
            // 遍历表格数据
            for (int i = 0, k = startRow; i < dateList.size(); i++, k++) {
                //TableData对象
                Object tableData = dateList.get(i);
                Object methodValue;
                // 从指定行开始写入（excel中的行默认从0开始）
                row = sheet.createRow(k);
                //设置单元格样式
                CellStyle style = wb.createCellStyle();
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平居中
                style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); //垂直居中
                style.setWrapText(true);    //自动换行
                Font font = wb.createFont();  //创建一个字体，用来保存字体样式
                font.setFontHeightInPoints((short) 10); //字体大小
                font.setFontName("微软雅黑");   //字体类型
                style.setFont(font);    //设置字体样式
                //用来记录插入列的下标
                int index = 0;
                //遍历表头---根据表头插入相应数据
                for (int j = 0; j < titles.size(); j++) {
                    //保存表头名称
                    String title;
                    Object titleItem = titles.get(j);
                    if (titleItem instanceof String) {  //判断表头类型是否是String,即判断该表头在对象中是否是String类型的属性
                        //调用get方法，获取返回值
                        methodValue = getFieldValue(titleItem.toString(), tableDataClass, tableData);
                        //创建
                        Cell cell1 = row.createCell(index);
                        //设置列值
                        cell1.setCellValue((String) methodValue);
                        //设置列样式
                        cell1.setCellStyle(style);

                        index++;
                    } else {    //该表头为Map类型,即该表头在对象中是对象类型的属性
                        Map<String, Object> map = (Map<String, Object>) titleItem;
                        for (Map.Entry<String, Object> mapItem : map.entrySet()) {
                            //获取tableData 属性对象的属性值
                            List<Object> value1 = (List<Object>) mapItem.getValue();
                            //调用get方法，获取返回值
                            methodValue = getFieldValue(mapItem.getKey(), tableDataClass, tableData);
                            //获取该属性
                            Field field1 = tableDataClass.getDeclaredField(mapItem.getKey());
                            //获取该属性对象所对应的类
                            Class<?> fieldClass = field1.getType().newInstance().getClass();
                            //遍历对象属性
                            for (Object o : value1) {
                                //调用get方法，获取返回值
                                Object fieldMethodValue = getFieldValue(o.toString(), fieldClass, methodValue);
                                Cell cell1 = row.createCell(index);
                                cell1.setCellValue(fieldMethodValue == null || fieldMethodValue == "" ? "无" : (String) fieldMethodValue);
                                cell1.setCellStyle(style);

                                index++;
                            }
                        }
                    }
                }
            }
            OutputStream outputStream = new FileOutputStream(excelUrl);
            wb.write(outputStream);
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            try {
                if (out != null) {
                    out.flush();
                    out.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * Description:
     *
     * @param title     对象中的属性名称
     * @param classType 该对象所对应的类
     * @param instance  对象实例
     * @Return: java.lang.Object  调用方法后的返回值
     * @Date: 2021/03/22 13:22
     */
    public static Object getFieldValue(String title, Class<?> classType, Object instance) throws Exception {
        Object methodValue;
        //获取该属性所对应的get方法
        String getMethodName = "get" + title.substring(0, 1).toUpperCase() + title.substring(1);
        //执行指定方法,获取方法返回值
        Method method = classType.getDeclaredMethod(getMethodName);
        methodValue = method.invoke(instance);
        return methodValue;
    }


    /**
     * Description: 获取Excel数据行数
     *
     * @param fileUrl  文件路径
     * @param sheetNum sheet数
     * @Return: int
     * @Date: 2021/03/22 14:34
     */
    public static int getRow(String fileUrl, int sheetNum) throws Exception {
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);

        Sheet sheet = workbook.getSheetAt(sheetNum - 1); //读取第几个工作表sheet
        int rowNum = sheet.getLastRowNum();//有多少行
        int num = 0;
        for (int i = 0; i < rowNum; i++) {
            if (sheet.getRow(i) == null)
                break;
            else
                num++;
        }
        return num;
    }


}
