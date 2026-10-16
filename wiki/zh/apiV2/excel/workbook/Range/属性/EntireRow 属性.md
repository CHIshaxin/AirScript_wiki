---
title: "EntireRow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/EntireRow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BqGNXrq3
---

# EntireRow (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示包含指定区域的整行（或多行）。只读。

## 示例

```javascript
/*本示例对包含活动单元格的行中的第一个单元格赋值。本示例必须在工作表上运行。*/
function test() {
    Application.ActiveCell.EntireRow.Cells.Item(1, 1).Value2 = 5
}
```

```javascript
/*本示例将 B2 单元格的值设置为6，并删除 A1 单元格所在的行*/
function test() {
    Range("B2").Value2 = 6
    Range("A1").EntireRow.Delete()
}
```
