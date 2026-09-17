---
title: "SqrtPi 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SqrtPi%20%E6%96%B9%E6%B3%95.html
---
# SqrtPi (方法)

返回某数与 pi 的乘积的平方根。

## 说明

如果 number < 0，则 SQRTPI 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 与 pi 相乘的数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SqrtPi 方法计算所给出数值与 pi 的乘积的平方根，并分别赋值给A3和A4单元格。*/
function test() {
    Range("A3").Value2 = WorksheetFunction.SqrtPi(2.3)
    Range("A4").Value2 = WorksheetFunction.SqrtPi(2)
}
```

```javascript
/*本示例为B1单元格赋值，使用 SqrtPi 方法计算B1单元格的值与 pi 的乘积的平方根。*/
function test() {
    Range("B1").Value2 = 3.6
    console.log(WorksheetFunction.SqrtPi(Range("B1").Value2))
}
```
