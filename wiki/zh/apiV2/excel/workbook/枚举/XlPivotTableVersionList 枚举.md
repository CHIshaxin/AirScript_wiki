---
title: "XlPivotTableVersionList 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlPivotTableVersionList%20%E6%9E%9A%E4%B8%BE.html
updated: 1789700916000
hash: BO-VIcA_
---

# XlPivotTableVersionList 枚举

---

指定数据透视表或数据透视表缓存的版本。创建特定版本的数据透视表可确保在 中创建的表的行为方式与它们在 ET 的相应版本中的行为方式相同。

| 注释 |
| --- |
| *xlPivotTableVersionCurrent* 仅出于向后兼容性的原因而包含在内。它不能与新的 **PivotCache** 和 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象一起使用。*xlPivotTableVersion11* 和 *xlPivotTableVersion10* 之间的行为并无差异。 |

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlPivotTableVersion2000** | 0 | ET 2000 |
| **xlPivotTableVersion10** | 1 | ET 2002 |
| **xlPivotTableVersion11** | 2 | ET 2003 |
| **xlPivotTableVersion12** | 3 | ET 2007 |
| **xlPivotTableVersion14** | 4 | ET 2010 |
| **xlPivotTableVersionCurrent** | -1 | 仅为向后兼容性而提供 |
