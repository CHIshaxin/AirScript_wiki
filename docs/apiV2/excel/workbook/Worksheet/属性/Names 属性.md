---
title: "Names 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Names%20%E5%B1%9E%E6%80%A7.html
---
# Names (属性)

返回一个 Names集合，它代表所有特定于工作表的名称（使用“WorksheetName!”前缀定义的名称）。 **[Names](./../../Names/Names 对象.html)** 对象，只读。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Names`。

## 示例

```javascript
/*本示例是将 Sheet1 中的 A1 单元格的名称定义为“myName”。*/
function test() {
    Worksheets.Item("Sheet1").Names.Add("myName", null, null, null, null, null, null, null, null, "=Sheet1!R1C1")
}
```

```javascript
/*本示例为A1单元格定义新名称后，遍历Names集合。*/
function test() {
    ActiveSheet.Names.Add("newName", null, null, null, null, null, null, null, null, "=Sheet1!R1C1")
    for (let i = 1; i <= ActiveSheet.Names.Count; i++) {
        ActiveSheet.Cells(1, i).Value2 = ActiveSheet.Names.Item(i).Name
    }
}
```
