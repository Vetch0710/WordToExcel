package com.wordToExcel.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * @author Vetch
 * @Description
 * @create 2021-03-18-13:53
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class IndexCalibre {
    //准则口径
    private String criterionCalibre="无";

    //业财口径
    private String fortuneCalibre="无";

}
