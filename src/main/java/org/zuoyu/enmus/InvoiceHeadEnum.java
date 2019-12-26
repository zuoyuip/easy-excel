package org.zuoyu.enmus;

import lombok.Getter;

/**
 * @author : zuoyu
 * @project : esay-excel
 * @description : 发票抬头
 * @date : 2019-12-26 10:20
 **/
@Getter
public enum InvoiceHeadEnum {
    /**
     * 个人
     */
    PERSONAL(0),

    /**
     * 企业
     */
    COMPANY(1);

    private Integer value;

    InvoiceHeadEnum(Integer value) {
        this.value = value;
    }
}
