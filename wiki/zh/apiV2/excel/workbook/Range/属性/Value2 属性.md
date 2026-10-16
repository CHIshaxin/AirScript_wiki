---
title: "Value2 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Value2%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: tjRPWdIZ
---

# Value2 (属性)

返回或设置单元格值。**Variant** 类型，可读写。

## 说明

该属性与 **Value** 属性的唯一区别是：**Value2** 属性不使用 **Currency** 和 **Date** 数据类型。可以通过使用 **Double** 数据类型，以浮点数形式返回这些数据类型格式的数值。

## 示例

```javascript
/*本示例使用 Value2 属性对两个单元格的值进行相加。*/
function test() {
    Application.Range("A1").Value2 = Range("B1").Value2 + Range("C1").Value2
}
```

```javascript
/*本示例在 A1 单元格设置公式 =SUM(B1:C1)，并显示 A1 单元格的值*/
function test() {
    Range("A1").Formula = "=SUM(B1:C1)"
    console.log(Range("A1").Value2)
}
```
