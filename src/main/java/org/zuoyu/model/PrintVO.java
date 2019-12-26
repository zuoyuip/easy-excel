package org.zuoyu.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;
import org.zuoyu.convert.InvoiceHeadEnumConverter;
import org.zuoyu.convert.InvoiceTypeEnumConverter;
import org.zuoyu.enmus.InvoiceHeadEnum;
import org.zuoyu.enmus.InvoiceTypeEnum;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author : zuoyu
 * @project : esay-excel
 * @description : 打印示例
 * @date : 2019-12-26 10:17
 **/
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class PrintVO {

    @ExcelProperty(value = "序号")
    private Integer number;

    @ExcelProperty(value = "订单编号")
    private String orderCode;

    @ExcelProperty(value = "用户电话")
    private String userMobile;

    @ExcelProperty(value = "支付时间")
    @DateTimeFormat(value = "yyyy年MM月dd日HH时mm分ss秒")
    @JsonFormat(pattern="yyyy-MM-dd HH:mm:ss")
    private Date createTime;

    @NumberFormat(value = "\u00A4#,###,###.00")
    @ExcelProperty(value = "支付金额")
    private BigDecimal amount;

    @ExcelProperty(value = "最晚开票时间")
    @DateTimeFormat(value = "yyyy年MM月dd日HH时mm分ss秒")
    @JsonFormat(pattern="yyyy-MM-dd HH:mm:ss")
    private Date invoiceTime;

    @ExcelProperty(value = "发票抬头", converter = InvoiceHeadEnumConverter.class)
    private InvoiceHeadEnum head;

    @ExcelProperty(value = "发票类型", converter = InvoiceTypeEnumConverter.class)
    private InvoiceTypeEnum type;

    @ExcelProperty(value = "发票内容")
    private String content;

    @ExcelProperty(value = "发票公司")
    private String company;

    @ExcelProperty(value = "单位税号")
    private String taxNumber;

    @ExcelProperty(value = "收货人电话")
    private String mobile;

    @ExcelProperty(value = "收货人邮件")
    private String email;

    @ExcelProperty(value = "收货人")
    private String consignee;

    @ExcelProperty(value = "收货人地址")
    private String address;
}
