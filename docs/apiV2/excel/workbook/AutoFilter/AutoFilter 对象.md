---
title: "AutoFilter 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AutoFilter/AutoFilter%20%E5%AF%B9%E8%B1%A1.html
---
# AutoFilter (对象)

代表对指定工作表的自动筛选。

## 说明

使用 **AutoFilter** 属性可返回 **[AutoFilter](./../AutoFilter/AutoFilter 对象.html)** 对象。使用 **Filters** 属性可返回由各个列筛选组成的集合。使用 **Range** 属性可返回代表整个筛选区域的 **[Range](./../Range/Range 对象.html)** 对象。

要为工作表创建 **[AutoFilter](./../AutoFilter/AutoFilter 对象.html)** 对象，必须手动打开或使用 **[Range](./../Range/Range 对象.html)** 对象的 **AutoFilter** 方法打开工作表上某个区域上的自动筛选功能。

| 注释 |
| --- |
| When using **AutoFilter** with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid. |
