package com.wordToExcel.tool;

import com.wordToExcel.entity.TableData;
//import com.wordToExcel.tool.ExcelTool;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 读取word文档中表格数据，支持doc、docx
 *
 * @author Fise19
 */
public class WordTools {



    /**
     * Description: 读取表格数据
     *
     * @param filePath   word文件路径
     * @param properties 存放word与对象属性对应关系的配置文件对象
     * @Return: java.util.List<com.wordToExcel.entity.TableData>    表格数据
     * @Date: 2021/03/22 13:51
     */
    public static List<TableData> tableInWord(String filePath, Properties properties) {
        try {
            FileInputStream in = new FileInputStream(filePath);//载入文档
            // 处理docx格式 即office2007以后版本
            if (filePath.toLowerCase().endsWith("docx")) {
                //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
                XWPFDocument xwpf = new XWPFDocument(in);//得到word文档的信息
                Iterator<XWPFTable> it = xwpf.getTablesIterator();//得到word中的表格
                //存放表格数据
                List<TableData> lists = new ArrayList<>();
                //获取类
                Class clazz = TableData.class;
                while (it.hasNext()) {

                    TableData tableDataItem = new TableData();
                    XWPFTable table = it.next();
                    List<XWPFTableRow> rows = table.getRows();
                    //读取每一行数据
                    for (int i = 1; i < rows.size(); i++) {
                        XWPFTableRow row = rows.get(i);
                        //读取每一列数据
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (int j = 0; j < cells.size(); j += 2) {
                            XWPFTableCell cell = cells.get(j);
                            //当前列为表格标题名称
                            String textKey = cell.getText();
                            if (j + 1 >= cells.size() || "".equals(textKey)) { //判断是否有下一列
                                continue;
                            }
                            //下一列为该标题所对应值
                            String textValue = "".equals(cells.get(j + 1).getText()) ? "无" : cells.get(j + 1).getText();
//                            String property = properties.getProperty(textKey) != null ? properties.getProperty(textKey) : textKey;
                            //获取该标题所对应对象中的属性值
                            String property = properties.getProperty(textKey);
                            if (property == null && !"定义说明".equals(textKey)) {  //过滤掉不需要的列值
                                continue;
                            }
                            if (!"定义说明".equals(textKey)) {
                                Field field = clazz.getDeclaredField(property);//获取属性
                                field.setAccessible(true);//允许操作
                                if (!(field.getType().isInstance("String"))) {
                                    Class<?> o = field.getType().newInstance().getClass();//获取属性对象所对应的类
                                    Object instance = field.getType().newInstance();   //属性对象实例
                                    Field[] declaredFields = o.getDeclaredFields();     //属性对象的全部属性
                                    //给属性对象的属性赋值
                                    WordTools.getValue(textValue, declaredFields, instance, properties);
                                    //给属性对象赋值
                                    field.set(tableDataItem, instance);

                                } else {
                                    //普通字符串类型则直接将值 赋值给属性
                                    field.set(tableDataItem, textValue);
                                }
                            } else {    //将计算方式作为特殊处理
                                Field field = clazz.getDeclaredField(properties.getProperty("计算方式"));
                                field.setAccessible(true);
                                int index = textValue.indexOf("计算方式");
                                if (index != -1) {
                                    textValue = textValue.substring(index);
                                } else {
                                    textValue = "无";
                                }
                                field.set(tableDataItem, textValue);
                            }
                        }
                    }
                    //过滤掉没有名称的表格，即没有数据的表格
                    if (!"无".equals(tableDataItem.getIndexName()))
                        lists.add(tableDataItem);

                }
                return lists;
            } else {
             /*   // 处理doc格式 即office2003版本
                POIFSFileSystem pfs = new POIFSFileSystem(in);
                HWPFDocument hwpf = new HWPFDocument(pfs);
                Range range = hwpf.getRange();//得到文档的读取范围
                TableIterator itpre = new TableIterator(range);
                ;//得到word中的表格
                int total = 0;
                while (itpre.hasNext()) {
                    itpre.next();
                    total += 1;
                }
                TableIterator it = new TableIterator(range);
                // 迭代文档中的表格
                // 如果有多个表格只读取需要的一个 set是设置需要读取的第几个表格，total是文件中表格的总数
                int set = orderNum;
                int num = set;
                for (int i = 0; i < set - 1; i++) {
                    it.hasNext();
                    it.next();
                }
                Map<String, Map<String, Object>> tableList = new HashMap<>();
                while (it.hasNext()) {
                    Table tb = (Table) it.next();
                    Map<String, Object> tableData = new HashMap<>();
                    System.out.println("这是第" + num + "个表的数据");
                    //迭代行，默认从0开始,可以依据需要设置i的值,改变起始行数，也可设置读取到那行，只需修改循环的判断条件即可
                    for (int i = 0; i < tb.numRows(); i++) {
                        TableRow tr = tb.getRow(i);
                        //迭代列，默认从0开始
                        for (int j = 0; j < tr.numCells(); j += 2) {
                            TableCell td = tr.getCell(j);//取得单元格
                            if (j + 1 >= tr.numCells()) {
                                continue;
                            }
                            //取得单元格的内容
//                            for (int k = 0; k < td.numParagraphs(); k++) {
                            Paragraph para = td.getParagraph(j);
                            String textKey = para.text();
                            System.out.println(td.getParagraph(j + 1).text());
                            String textValue = "111";
//                            String textValue = tr.getCell(j+1).getParagraph(j+1).text();

                            //去除后面的特殊符号
//                                if (null != s && !"".equals(s)) {
//                                    s = s.substring(0, s.length() - 1);
//                                }
                            if ("定义说明".equals(textKey)) {
                                int index = textValue.indexOf("计算方式");
                                if (index != -1) {
                                    textValue = textValue.substring(index);
                                } else {
                                    textValue = "";
                                }
                                tableData.put("计算方式", textValue.length() == 0 ? "无" : textValue);

                            } else {
                                if ("指标口径".equals(textKey)) {
                                    Map<String, String> value = WordReaderUtil2.getValue(textValue, new ArrayList<>(Arrays.asList("准则口径", "业财口径")), properties);
                                    tableData.put(textKey, value);


                                } else if ("指标用途".equals(textKey)) {
                                    Map<String, String> value = WordReaderUtil2.getValue(textValue, new ArrayList<>(Arrays.asList("内部经营", "外部报送")), properties);
                                    tableData.put(textKey, value);

                                } else {
                                    tableData.put(textKey, textValue.length() == 0 ? "无" : textValue);
                                }
                            }
//                            tableData.put(textKey, cells.get(j + 1).getText());

                        }

//                        System.out.println(tableData);
                    }
                    tableList.put(tableData.get("名称").toString(), tableData);
//                     过滤多余的表格
//                    while (num < total) {
//                        it.hasNext();
//                        it.next();
//                        num += 1;
//                    }
                    num++;
//                    }
////                                rowList.add(s);
////                                System.out.print(s + "[" + i + "," + j + "]" + "\t");
////                            }
////                        }
//                        tableList.add(rowList);
//                        System.out.println();
//                    }
                    // 过滤多余的表格
//                    while (num < total) {
//                        it.hasNext();
//                        it.next();
//                        num += 1;
//                    }
                }
                return null;*/
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }


    /**
     * Description:  当表格标题所对应对象的属性为对象时，设置并获取对象属性
     *
     * @param text           属性对象名称
     * @param declaredFields 该属性名称所对应的属性
     * @param instance       该属性对象实例
     * @param properties     存放匹配关系的配置文件
     * @Return: void
     * @Date: 2021/03/22 13:52
     */
    public static void getValue(String text, Field[] declaredFields, Object instance, Properties properties) throws Exception {

        String mid = text.replace(" ", "");

        String property, textKey, value;
        if (declaredFields.length <= 0) {
            throw new Exception("查询列表为空");
        }

        for (int i = 0; i < declaredFields.length; i++) {
            //允许操作私有属性
            declaredFields[i].setAccessible(true);
            //获取对象属性值
            textKey = declaredFields[i].getName();
            //获取对象属性 对应  word表格中的值
            property = properties.getProperty(textKey);
            //获取当前属性出现的位置
            int index1 = mid.indexOf(property);
            //保存字符串截取的终止位置
            int index2 = 0;
            if (index1 == -1) { //在表格中未找到当前属性，则设置为  “无”
                declaredFields[i].set(instance, "无");
                continue;
            } else if (i + 1 >= declaredFields.length) {    //若当前为最后一个属性，则将终止位置设为字符串的长度
                index2 = mid.length();
            } else {    //对象中有多个属性，则获取下一个属性出现的位置作为终止位置
                index2 = mid.indexOf(properties.getProperty(declaredFields[i + 1].getName()));
            }
            //通过截取字符串 获取当前对象属性 所对应 在表格中的属性值
            value = mid.substring(index1 + property.length() + 1, index2);
            declaredFields[i].set(instance, value);
        }


    }


}