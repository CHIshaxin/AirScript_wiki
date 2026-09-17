---
title: "SeriesSum 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SeriesSum%20%E6%96%B9%E6%B3%95.html
---
# SeriesSum (方法)

返回基于以下公式的幂级数之和：SERIES(x,n,m,a) = a₁xⁿ + a₂x⁽ⁿ⁺ᵐ⁾ + a₃x⁽ⁿ⁺²ᵐ⁾ + ... + aᵢx⁽ⁿ⁺⁽ⁱ⁻¹⁾ᵐ⁾。

## 说明

如果任一参数为非数值型，则 SERIESSUM 返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | X - 幂级数的输入值。 |
| Arg2 | any | 必填 | N - x 的首项乘幂。 |
| Arg3 | any | 必填 | M - 级数中每一项的乘幂 n 的步长增加值。 |
| Arg4 | any | 必填 | Coefficients - 一系列与 x 各级乘幂相乘的系数。coefficients 中值的数目决定了幂级数的项数。例如，如果 coefficients 中有三个值，幂级数中将有三项。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SeriesSum 方法计算幂级数之和，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.SeriesSum(-2.5, -1, -2, -5.3)
    Range("B2").Value2 = WorksheetFunction.SeriesSum(2, 2, 3, 4)
}
```

```javascript
/*本示例为A1等单元格赋值，使用 SeriesSum 方法计算幂级数之和。*/
function test() {
    Range("A1").Value2 = 3
    Range("B1").Value2 = 1.5
    Range("C1").Value2 = 2
    Range("D1").Value2 = 5.2
    console.log(WorksheetFunction.SeriesSum(Range("A1").Value2, Range("B1").Value2, Range("C1").Value2, Range("D1").Value2))
}
```
