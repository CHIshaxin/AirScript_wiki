---
title: "Precedents (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Precedents%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: nQj-X72Q
---

# Precedents (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示单元格的所有引用单元格。如果有多个引用单元格，则可以是一个多重选择（ **[Range](./../../Range/Range 对象.html)** 对象的并集）。只读。

## 示例

```javascript
/*本示例选定 sheet1 中单元格 A1 的引用单元格。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").Precedents.Select()
}
```

```javascript
/*本示例将 C3 单元格的公式设置为 =SUM(E3:G5)，显示 C3 单元格引用区域的地址*/
function test() {
    Range("C3").Formula = "=SUM(E3:G5)"
    console.log(Range("C3").Precedents.Address())
}
```
