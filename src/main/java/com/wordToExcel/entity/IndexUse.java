package com.wordToExcel.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * @author Vetch
 * @Description
 * @create 2021-03-18-13:55
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class IndexUse {
    //指标用途----内部经营
    private String internalManagement="无";

    //指标用途----外部报送
    private String externalSubmitted="无";
}
