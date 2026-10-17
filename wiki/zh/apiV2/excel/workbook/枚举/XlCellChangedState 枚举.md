---
title: "XlCellChangedState 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlCellChangedState%20%E6%9E%9A%E4%B8%BE.html
updated: 1789700916000
hash: i3-zq5E-
---

# XlCellChangedState 枚举

---

指定自创建数据透视表以来，或上次执行提交操作以来，数据透视表值单元格是否经过了编辑或重新计算。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlCellChangeApplied** | 3 | 单元格中的值已经过编辑或重新计算，并且这些更改已应用于数据源。（只应用具有 OLAP 数据源的数据透视表） |
| **xlCellChanged** | 2 | 单元格中的值已经过编辑或重新计算。 |
| **xlCellNotChanged** | 1 | 单元格中的值未经过编辑或重新计算。 |
