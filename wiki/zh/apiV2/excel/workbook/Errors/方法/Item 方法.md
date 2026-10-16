---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Errors/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BgJuFRtq
---

# Item (方法)

返回 **[Error](./../../Error/Error 对象.html)** 对象的单个成员。

## 说明

*Index* 可以是以下常量之一。

| **xlEvaluateToError**：单元格计算为错误值。 |
| --- |
| **xlTextDate**：单元格包含用 2 位数表示年份的文本日期。 |
| **xlNumberAsText**：单元格包含以文本形式存储的数字。 |
| **xlInconsistentFormula**：单元格包含一个区域中不一致的公式。 |
| **xlOmittedCells**：单元格包含一个忽略了区域中某个单元格的公式。 |
| **xlUnlockedFormulaCells**：取消锁定的单元格包含一个公式。 |
| **xlEmptyCellReferences**：单元格包含一个引用空单元格的公式。 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 成员的索引。 |
