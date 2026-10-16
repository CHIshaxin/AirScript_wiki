---
title: "Width (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C7P7pgU_
---

# Width (属性)

返回一个 **Variant** 值，它代表区域的宽度（以单位表示）。

## 示例

```javascript
/*本示例将 A1 单元格的宽度乘以 2，然后显示 A1 单元格的宽度*/
function test() {
    Range("A1").ColumnWidth *= 2
    console.log(Range("A1").Width)
}
```

```javascript
/*本示例显示 A1:B1 单元格区域的宽度*/
function test() {
    console.log("A1:B1区域的总宽度为：" + Range("A1:B1").Width)
}
```
