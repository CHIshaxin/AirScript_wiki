---
title: "SumX2MY2 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SumX2MY2%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DOwgnFkW
---

# SumX2MY2 (方法)

返回两个数组中对应值的平方差之和。

## 说明

- 参数应为数字，或者是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果 array_x 和 array_y 中值的数目不同，则 SUMX2MY2 将返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array_x - 第一个数值数组或区域。 |
| Arg2 | any | 必填 | Array_y - 第二个数值数组或区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SumX2MY2 方法计算两个数组中对应值的平方差之和，并赋值给D2单元格。*/
function test() {
    let arr1 = [10, -10.3, 30.2]
    let arr2 = [10.3, 3.2, 6.2]
    Range("D2").Value2 = WorksheetFunction.SumX2MY2(arr1, arr2)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 SumX2MY2 方法计算两个数组中对应值的平方差之和。*/
function test() {
    Range("A1").Value2 = 23.3
    Range("A2").Value2 = 20.03
    Range("A3").Value2 = 20.6
    Range("A4").Value2 = 21
    Range("B1").Value2 = -5
    Range("B2").Value2 = 2
    Range("B3").Value2 = 6
    Range("B4").Value2 = 3
    console.log(WorksheetFunction.SumX2MY2(Range("A1:A4"), Range("B1:B4")))
}
```
