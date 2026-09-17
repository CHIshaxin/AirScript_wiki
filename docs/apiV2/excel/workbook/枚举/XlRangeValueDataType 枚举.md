---
title: "XlRangeValueDataType 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlRangeValueDataType%20%E6%9E%9A%E4%B8%BE.html
---
# XlRangeValueDataType 枚举

---

指定区域值数据类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlRangeValueDefault** | 10 | 默认值。如果指定的 **[Range](./../Range/Range 对象.html)** 对象为空，则返回值 Empty（可用 IsEmpty 函数测试这种情况）。如果 **[Range](./../Range/Range 对象.html)** 对象包含多个单元格，则返回值的数组（可用 IsArray 函数测试这种情况）。 |
| **xlRangeValueMSPersistXML** | 12 | 以 XML 格式返回指定的 **[Range](./../Range/Range 对象.html)** 对象的记录集表示形式。 |
| **xlRangeValueXMLSpreadsheet** | 11 | 以 XML 电子表格格式返回指定的 **[Range](./../Range/Range 对象.html)** 对象的值、格式、公式和名称。 |
