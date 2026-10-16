---
title: "BesselI 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/BesselI%20%E6%96%B9%E6%B3%95.html
---
# BesselI (方法)

返回修正的 Bessel 函数，它等效于计算纯虚数参数值的 Bessel 函数。

## 说明

- 如果 x 是非数值，BesselI 将返回错误值 #VALUE!。
- 如果 n 是非数值，BesselI 将生成一个错误值。
- 如果 n < 0，BesselI 将生成一个错误值。
- 变量 x 的 n 阶修正 Bessel 函数为：

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 要计算的函数值。 |
| Arg2 | any | 必填 | Bessel 函数的阶。如果 n 不是整数，则将被截尾取整。 |

## 返回值

Double

## 示例

```javascript
/*本示例为D5等单元格分别赋值，并使用BesselI方法返回修正Bessel函数值。*/
function test() {
    Range("D5").Value2 = 1
    Range("E5").Value2 = 2
    Range("F5").Value2 = 4.5
    Range("G5").Value2 = 7
    let besseli1 = Application.WorksheetFunction.BesselI(Range("G5").Value2, Range("D5").Value2)
    let besseli2 = Application.WorksheetFunction.BesselI(Range("F5").Value2, Range("E5").Value2)
    console.log(besseli1)
    console.log(besseli2)
}
```

```javascript
/*本示例使用BesselI方法返回修正Bessel函数值，并将结果赋值于C4等单元格。*/
function test() {
    Range("C4").Value2 = Application.WorksheetFunction.BesselI(-1.5, 1)
    Range("C5").Value2 = Application.WorksheetFunction.BesselI(11, 3)
    Range("C6").Value2 = Application.WorksheetFunction.BesselI(20.25, 4)
}
```
