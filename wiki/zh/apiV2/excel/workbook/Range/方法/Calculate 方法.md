---
title: "Calculate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Calculate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C-5R5H1m
---

# Calculate (方法)

计算所有打开的工作簿、工作簿中的某张特定工作表或工作表指定区域中的单元格，如下表所示。

## 说明

| 要计算 | 依照此示例 |
| --- | --- |
| 所有打开的工作簿 | Application.Calculate()（或只是Calculate()） |
| 指定工作表 | Worksheets.Item(1).Calculate() |
| 指定区域 | Worksheets.Item(1).Rows.Item(2).Calculate() |

## 返回值

Variant

## 示例

```javascript
/*此示例计算 Sheet1 上已用区域中 A 列、B 列和 C 列的公式。*/
function test() {
    Application.Worksheets.Item("Sheet1").UsedRange.Columns.Item("A:C").Calculate()
}
```

```javascript
/*本示例将 B3 单元格的公式设置为 =RANDBETWEEN(12,88)，并重新计算 B3 单元格的公式*/
function test() {
    Range("B3").Formula = "=RANDBETWEEN(12,88)"
    Range("B3").Calculate()
}
```
