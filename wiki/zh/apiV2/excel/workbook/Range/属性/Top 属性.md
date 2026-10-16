---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DYGn2edY
---

# Top (属性)

返回或设置一个 **Variant** 值，它代表行 1 上边缘到区域上边缘的距离（以磅为单位）。

## 说明

如果该区域不连续，则使用第一块区域。如果该区域有若干行，则使用区域中的顶端行（行号最小的行）。

## 示例

```javascript
/*本示例显示 A2 单元格上边缘到第一行上边缘的距离*/
function test() {
    console.log(Range("A2").Top)
}
```

```javascript
/*本示例将 A1 单元格所在行的行高设置为21，并显示 A2 单元格上边缘到第一行上边缘的距离*/
function test() {
    Range("A1").Rows.RowHeight = 21
    console.log(Range("A2").Top)
}
```
