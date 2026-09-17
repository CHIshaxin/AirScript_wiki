---
title: "Resize 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Resize%20%E6%96%B9%E6%B3%95.html
---
# Resize (方法)

调整指定区域的大小。返回 **[Range](./../../Range/Range 对象.html)** 对象，该对象代表调整后的区域。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowSize | any | 可选 | 新区域中的行数。如果省略该参数，则该区域中的行数保持不变。 |
| ColumnSize | any | 可选 | 新区域中的列数。如果省略该参数。则该区域中的列数保持不变。 |

## 示例

```javascript
/*本示例调整 Sheet1 中选定区域的大小，使之增加一行和一列。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    numRows = Application.Selection.Rows.Count
    numColumns = Application.Selection.Columns.Count
    Application.Selection.Resize(numRows + 1, numColumns + 1).Select()
}
```

```javascript
/*本示例假定在 Sheet1 中有一个包含标题行的表。本示例选定该表，但不选定标题行。运行本示例之前，活动单元格必须处于该表中。*/
function test() {
    let tbl = Application.ActiveCell.CurrentRegion
    tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, tbl.Columns.Count).Select()
}
```
