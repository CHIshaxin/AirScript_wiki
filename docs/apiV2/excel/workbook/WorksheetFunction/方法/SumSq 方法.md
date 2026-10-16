---
title: "SumSq 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SumSq%20%E6%96%B9%E6%B3%95.html
---
# SumSq (方法)

返回参数的平方之和。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的数字、逻辑值和数字的文本表示也包括在内。
- 如果参数为数组或引用，则只有该数组或引用中的数字将被计算在内。数组或引用中的空单元格、逻辑值、文本或错误值将被忽略。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number1, number2, ... - 1 到 30 个需要求平方和的参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SumSq 方法计算参数的平方之和，并分别赋值给C2和C3单元格。*/
function test() {
    let arr1 = [10, -10.3, 30.2]
    Range("C2").Value2 = WorksheetFunction.SumSq(arr1)
    Range("C3").Value2 = WorksheetFunction.SumSq(1, 2, 3, 4, 5, 6, 7, 8, 9)
}
```

```javascript
/*本示例为B1等单元格赋值，使用 SumSq 方法计算参数的平方之和。*/
function test() {
    Range("B1").Value2 = -5
    Range("B2").Value2 = 2
    Range("B3").Value2 = 6
    Range("B4").Value2 = 3
    console.log(WorksheetFunction.SumSq(Range("B1:B4")))
}
```
