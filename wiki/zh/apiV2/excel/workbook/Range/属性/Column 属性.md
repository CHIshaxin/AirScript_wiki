---
title: "Column (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Column%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bcm4fZqr
---

# Column (属性)

返回指定区域中第一块中的第一列的列号。Long 类型，只读。

## 说明

A 列返回 1，B 列返回 2，依此类推。

若要返回区域中最后一列的列号，请使用下列语句。

myRange.Columns(myRange.Columns.Count).Column

## 示例

```javascript
/*本示例将工作表 Sheet1 上每隔一列的列宽设置为 4 磅。*/
function test() {
    let col = Application.Worksheets.Item("Sheet1").Columns
    for (let i = 1; i <= col.Count; i++) {
        if ((col.Item(i).Column) % 2 == 0) {
            col.Item(i).ColumnWidth = 4
        }
    }
}
```

```javascript
/*本示例将工作表 Sheet1 上第一列的文字加粗。*/
function test() {
    Application.Worksheets.Item("Sheet1").Columns.Item(1).Font.Bold = true
}
```
