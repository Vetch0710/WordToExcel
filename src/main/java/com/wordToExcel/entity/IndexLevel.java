package com.wordToExcel.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * @author Vetch
 * @Description
 * @create 2021-03-18-13:49
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class IndexLevel {
    //一类
    private String indexLevel1;

    //二类
    private String indexLevel2;

    //三类
    private String indexLevel3;
}
