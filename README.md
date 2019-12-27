# easy-excel
基于easyExcel的Excel表格导出工具

```
easy-excel
│ 
├─convert  转换器  
│    ├─InvoiceHeadEnumConverter.java  对发票抬头枚举的转换
│    └─InvoiceTypeEnumConverter.java  对发票类型枚举的转换
│    
├─enmus  枚举类
│  
├─model  模型
│  └─PrintVO  打印的数据模型（内有easyExcel的注解使用案例）
│    
├─util 工具包
│    ├─ExcelUtil.java  核心工具包
│    └─JsonUtil.java  JSON工具包（在此案例里仅用于数据转换）
│
├──resources 
│     │
│     └─data.json  打印列表数据源
│
```
