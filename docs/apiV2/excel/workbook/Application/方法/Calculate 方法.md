---
title: "Calculate 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/Calculate%20%E6%96%B9%E6%B3%95.html
---
# Calculate (方法)

计算所有打开的工作簿、工作簿中的某张特定工作表或工作表指定区域中的单元格，如下表所示。

## 说明

| 要计算 | 依照此示例 |
| --- | --- |
| 所有打开的工作簿 | Application.Calculate（或只是 Calculate） |
| 指定工作表 | Worksheets(1).Calculate |
| 指定区域 | Worksheets(1).Rows(2).Calculate |

## 示例

```javascript
/*此示例计算 Sheet1 上已用区域中 A 列、B 列和 C 列的公式。*/
function test() {
    Worksheets.Item("Sheet1").UsedRange.Columns("A:C").Calculate()
}
```

```javascript
/*此示例计算所有打开的工作簿。*/
function test() {
    Application.Calculate()
}
```
