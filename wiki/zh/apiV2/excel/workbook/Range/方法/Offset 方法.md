---
title: "Offset (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Offset%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DOxxxAyH
---

# Offset (方法)

返回 **[Range](./../../Range/Range 对象.html)** 对象，它代表位于指定单元格区域的一定的偏移量位置上的区域。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowOffset | any | 可选 | 区域偏移的行数（正数、负数或 0（零））。正数表示向下偏移，负数表示向上偏移。默认值是 0。 |
| ColumnOffset | any | 可选 | 区域偏移的列数（正数、负数或 0（零））。正数表示向右偏移，负数表示向左偏移。默认值是 0。 |

## 示例

```javascript
/*此示例激活 Sheet1 上活动单元格向右偏移三列、向下偏移三行处的单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Offset(3, 3).Activate()
}
```

```javascript
/*此示例假设 Sheet1 中包含一个具有标题行的表格。此示例先选定该表格，但并不选择行首。运行此示例之前，活动单元格必须位于表格中。*/
function test() {
    let tbl = Application.ActiveCell.CurrentRegion
    tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, tbl.Columns.Count).Select()
}
```
