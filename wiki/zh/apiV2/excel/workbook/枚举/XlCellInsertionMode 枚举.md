---
title: "XlCellInsertionMode 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlCellInsertionMode%20%E6%9E%9A%E4%B8%BE.html
updated: 1789700916000
hash: DiiPJkzA
---

# XlCellInsertionMode 枚举

---

指定在指定工作表中添加或删除行的方式，以符合查询返回的记录集的行数。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlInsertDeleteCells** | 1 | 插入或者删除部分行以符合新记录集所需要的实际行数。 |
| **xlInsertEntireRows** | 2 | 必要时插入所有行以允许溢出。不从工作表删除单元格或行。 |
| **xlOverwriteCells** | 0 | 不向工作表添加新的单元格或行。如果溢出则覆盖周围单元格中的数据。 |
