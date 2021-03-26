package com.wordToExcel.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;



/**
 * @author Vetch
 * @Description
 * @create 2021-03-18-13:48
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class TableData {

    // 指标分类
    private IndexLevel indexLevel = new IndexLevel();

    //指标分级
    private String indexGrade = "无";

    //指标编码
    private String indexEncoding = "无";

    //指标名称
    private String indexName = "无";

    //填报部门
    private String Department = "无";

    //指标口径
    private IndexCalibre indexCalibre = new IndexCalibre();

    //指标定义
    private String definition = "无";

    //计算方式
    private String calculation = "无";

    //计量单位
    private String unit = "无";

    //指标用途
    private IndexUse indexUse = new IndexUse();

    //适用范围
    private String scope = "无";

    //存量指标取数路径
    private String stockPath = "无";
}
