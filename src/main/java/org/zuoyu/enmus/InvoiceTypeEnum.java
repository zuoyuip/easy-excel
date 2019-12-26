package org.zuoyu.enmus;

import lombok.Getter;

/**
 * @author : zuoyu
 * @project : esay-excel
 * @description : 发票类型
 * @date : 2019-12-26 10:23
 **/
@Getter
public enum InvoiceTypeEnum {
    /**
     * 电子发票
     */
    ELECTRONIC(0),

    /**
     * 纸质发票
     */
    PAPER(1);

    private Integer value;

    InvoiceTypeEnum(Integer value) {
        this.value = value;
    }
}
