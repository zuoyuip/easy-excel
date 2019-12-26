package org.zuoyu.convert;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import org.zuoyu.enmus.InvoiceHeadEnum;

/**
 * @author : zuoyu
 * @project : esay-excel
 * @description : 发票抬头枚举类型转换
 * @date : 2019-12-26 15:07
 **/
public class InvoiceHeadEnumConverter implements Converter<InvoiceHeadEnum> {
    @Override
    public Class supportJavaTypeKey() {
        return InvoiceHeadEnum.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 读
     */
    @Override
    public InvoiceHeadEnum convertToJavaData(CellData cellData, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        return null;
    }

    /**
     * 写
     */
    @Override
    public CellData convertToExcelData(InvoiceHeadEnum invoiceHeadEnum, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        if (invoiceHeadEnum.equals(InvoiceHeadEnum.PERSONAL)){
            return new CellData("个人");
        }
        if (invoiceHeadEnum.equals(InvoiceHeadEnum.COMPANY)){
            return new CellData("企业");
        }
        return new CellData("未知");
    }
}
