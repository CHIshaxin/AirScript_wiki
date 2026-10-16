---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: vRgkjDRm
---

# Orientation (属性)

返回或设置一个 **Variant** 值，它代表文本方向。

## 说明

此属性的值可设为 –90 到 90 度之间的整数值或以下常量之一：**xlDownward**、**xlHorizontal**、**xlUpward** 、**xlVertical**

## 示例

```javascript
/*将 A1 单元格的文本方向设置为向上，A2 单元格的文本方向设置为居中向下*/
function test() {
    Range("A1").Orientation = xlUpward
    Range("A2").Orientation = xlVertical
}
```

```javascript
/*本示例将 C3:E5 单元格区域的文本倾斜 -30 度*/
function test() {
    Range("C3:E5").Orientation = -30
}
```
