package org.zuoyu.convert;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import org.zuoyu.enmus.InvoiceTypeEnum;

/**
 * @author : zuoyu
 * @project : esay-excel
 * @description : 发票类型枚举类型转换
 * @date : 2019-12-26 15:14
 **/
public class InvoiceTypeEnumConverter implements Converter<InvoiceTypeEnum> {
    @Override
    public Class supportJavaTypeKey() {
        return InvoiceTypeEnum.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 读
     */
    @Override
    public InvoiceTypeEnum convertToJavaData(CellData cellData, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        return null;
    }

    /**
     * 写
     */
    @Override
    public CellData convertToExcelData(InvoiceTypeEnum invoiceTypeEnum, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        if (invoiceTypeEnum.equals(InvoiceTypeEnum.ELECTRONIC)){
            return new CellData("电子发票");
        }
        if (invoiceTypeEnum.equals(InvoiceTypeEnum.PAPER)){
            return new CellData("纸质发票");
        }
        return new CellData("未知");
    }
}
