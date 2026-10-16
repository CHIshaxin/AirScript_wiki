---
title: "EntireColumn (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/EntireColumn%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BPBDqB1j
---

# EntireColumn (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示包含指定区域的整列（或多列）。只读。

## 示例

```javascript
/*本示例对包含活动单元格的列中的第一个单元格赋值。本示例必须在工作表上运行。*/
function test() {
    Application.ActiveCell.EntireColumn.Cells.Item(1, 1).Value2 = 5
}
```

```javascript
/*本示例设置 B2 的值为 b2，将 A1 单元格所在列删除*/
function test() {
    Range("B2").Value2 = 'b2'
    Range("A1").EntireColumn.Delete()
}
```
