---
title: "SumProduct 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SumProduct%20%E6%96%B9%E6%B3%95.html
---
# SumProduct (方法)

在给定的数组中，将数组中对应的元素相乘，并返回乘积之和。

## 说明

- 数组参数必须具有相同的维数，否则，SUMPRODUCT 将返回错误值 #VALUE!。
- SUMPRODUCT 将非数字型的数组项作为 0 处理。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Array1, array2, array3, ... - 2 到 30 个数组，其相应元素需要相乘并求和。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SumProduct 方法在给定的数组中，将数组中对应的元素相乘，计算乘积之和，并分别赋值给B3和B4单元格。*/
function test() {
    Range("B3").Value2 = WorksheetFunction.SumProduct(-1, 2, -3, 7, 50.8, 9.89, 10.2)
    Range("B4").Value2 = WorksheetFunction.SumProduct(1, 2, 3, 7, 8, 9, 10.2)
}
```

```javascript
/*本示例为B1等单元格赋值，使用 SumProduct 方法在给定的数组中，将数组中对应的元素相乘，并计算乘积之和。*/
function test() {
    Range("B1").Value2 = 5
    Range("B2").Value2 = 2
    Range("B3").Value2 = 9
    Range("B4").Value2 = 7
    Range("C1").Value2 = 3
    Range("C2").Value2 = 6
    Range("C3").Value2 = 0.3
    Range("C4").Value2 = 1
    console.log(WorksheetFunction.SumProduct(Range("B1:B4"), Range("C1:C4")))
}
```
