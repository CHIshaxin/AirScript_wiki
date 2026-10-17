---
title: "Gcd (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Gcd%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: ogt3U7or
---

# Gcd (方法)

返回两个或多个整数的最大公约数，最大公约数是能分别将 Number1 和 Number2 除尽的最大整数。

## 说明

- 如果任一参数为非数值型，则 GCD 将返回错误值 #VALUE!。
- 如果任一参数小于零，则 GCD 将返回错误值 #NUM!。
- 任何值都能被 1 整除。
- 素数只能被其本身和 1 整除。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2, ... - 为 1 到 29 个值，如果任一值不是整数，则将被截尾取整。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Gcd 方法计算两个或多个整数的最大公约数，并分别赋值给D1和D3单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.Gcd(2, 6)
    Range("D3").Value2 = WorksheetFunction.Gcd(100, 20, 50)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Gcd 方法计算A1等单元格的最大公约数。*/
function test() {
    Range("A1").Value2 = 9
    Range("A2").Value2 = 15
    Range("A3").Value2 = 30
    console.log(WorksheetFunction.Gcd(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
