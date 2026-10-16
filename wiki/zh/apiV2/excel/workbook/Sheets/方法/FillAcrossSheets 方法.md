---
title: "FillAcrossSheets (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E6%96%B9%E6%B3%95/FillAcrossSheets%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: -APSCnzF
---

# FillAcrossSheets (方法)

将单元格区域复制到集合中所有其他工作表的同一位置。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Range | Range | 必填 | 要填充到集合中所有工作表上的单元格区域。该区域必须来自集合中的某个工作表。 |

## 示例

```javascript
/*本示例用工作表 Sheet1 上 A1:C5 单元格区域的内容填充工作表 Sheet5 和 Sheet7 上的相同区域。*/
function test() {
    let names = ["Sheet1", "Sheet5", "Sheet7"]
    Application.Sheets.Item(names).FillAcrossSheets(Application.Worksheets.Item("Sheet1").Range("A1:C5"))
}
```

```javascript
/*本示例用工作表 Sheet2 上 A1:B2 单元格区域的内容及格式填充工作表 Sheet3 和 Sheet4 上的相同区域。*/
function test() {
    let names = ["Sheet2", "Sheet3", "Sheet4"]
    Application.Sheets.Item(names).FillAcrossSheets(Application.Worksheets.Item("Sheet2").Range("A1:B2"), xlFillWithAll)
}
```
