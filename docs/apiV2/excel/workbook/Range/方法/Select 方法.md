---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 说明

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用 **Activate** 方法。

## 示例

```javascript
/*本示例将 A3 单元格选中*/
function test() {
    Range("A3").Select()
}
```

```javascript
/*本示例将 B1 单元格的值设置为 b1，选中 A1:C3 单元格区域，并清除选中区域的内容*/
function test() {
    Range("B1").Value2 = 'b1'
    Range("A1:C3").Select()
    Selection.Clear()
}
```
