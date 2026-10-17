---
title: "BesselY (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/BesselY%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 4QrMWFcE
---

# BesselY (方法)

返回 Bessel 函数，该函数也称为 Weber 函数或 Neumann 函数。

## 说明

- 如果 x 为非数值型，则 BesselY 将生成一个错误值。
- 如果 n 为非数值型，则 BesselY 将生成一个错误值。
- 如果 n < 0，则 BesselY 将生成一个错误值。
- 变量 x 的 n 阶 Bessel 函数为：

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 计算的函数值。 |
| Arg2 | any | 必填 | 函数的阶。如果 n 不是整数，则将被截尾取整。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A2等单元格分别赋值，并使用BesselY方法计算并显示结果。*/
function test() {
    Range("A2").Value2 = 1
    Range("B2").Value2 = 2
    Range("C2").Value2 = 2.5
    Range("D2").Value2 = 10
    let bessely1 = Application.WorksheetFunction.BesselY(Range("C2").Value2, Range("B2").Value2)
    let bessely2 = Application.WorksheetFunction.BesselY(Range("D2").Value2, Range("A2").Value2)
    console.log(bessely1)
    console.log(bessely2)
}
```

```javascript
/*本示例使用BesselY方法计算，并将结果赋值于B3等单元格。*/
function test() {
    Range("B3").Value2 = Application.WorksheetFunction.BesselY(3, 3)
    Range("B4").Value2 = Application.WorksheetFunction.BesselY(23.5, 4)
    Range("B5").Value2 = Application.WorksheetFunction.BesselY(40.12, 1)
}
```
