---
title: "CurrentRegion 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/CurrentRegion%20%E5%B1%9E%E6%80%A7.html
---
# CurrentRegion (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示当前区域。当前区域是以空行与空列的组合为边界的区域。只读。

## 说明

该属性对于许多自动展开选择以包括整个当前区域的操作很有用，如 **AutoFormat** 方法。

该属性不能用于被保护的工作表。

## 示例

```javascript
/*本示例选定工作表 Sheet1 上的当前区域。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.CurrentRegion.Select()
}
```

```javascript
/*本示例假定在 Sheet1 中有一个包含标题行的表。本示例选定该表，但不选定标题行。运行本示例之前，活动单元格必须处于该表中。*/
function test() {
    let tbl = Application.ActiveCell.CurrentRegion
    tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, tbl.Columns.Count).Select()
}
```
