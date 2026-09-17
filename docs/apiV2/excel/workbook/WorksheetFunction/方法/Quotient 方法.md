---
title: "Quotient 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Quotient%20%E6%96%B9%E6%B3%95.html
---
# Quotient (方法)

返回商的整数部分，该函数可用于舍掉商的小数部分。

## 说明

如果两个参数中的任意一个为非数值型，则 QUOTIENT 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Numerator - 被除数。 |
| Arg2 | any | 必填 | Denominator - 除数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Quotient 方法分别计算100除以3和12除以5的整数部分，并分别赋值给D2和D3单元格。*/
function test() {
    Range("D2").Value2 = WorksheetFunction.Quotient(100, 3)
    Range("D3").Value2 = WorksheetFunction.Quotient(12, 5)
}
```

```javascript
/*本示例为B1等单元格赋值，使用 Quotient 方法获取B1单元格中数值除以B2单元格中数值的整数部分。*/
function test() {
    Range("B1").Value2 = 10
    Range("B2").Value2 = 4
    console.log(WorksheetFunction.Quotient(Range("B1").Value2, Range("B2").Value2))
}
```
