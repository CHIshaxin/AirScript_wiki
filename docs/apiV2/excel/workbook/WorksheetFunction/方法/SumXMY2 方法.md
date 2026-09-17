---
title: "SumXMY2 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SumXMY2%20%E6%96%B9%E6%B3%95.html
---
# SumXMY2 (方法)

返回两个数组中对应值之差的平方和。

## 说明

- 参数应为数字，或者是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果 array_x 和 array_y 中值的数目不同，则 SUMXMY2 将返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array_x - 第一个数值数组或区域。 |
| Arg2 | any | 必填 | Array_y - 第二个数值数组或区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SumXMY2 方法计算两个数组中对应值之差的平方和，并赋值给A1单元格。*/
function test() {
    let arr1 = [1, 2, 3]
    let arr2 = [10.3, -31.2, 60.2]
    Range("A1").Value2 = WorksheetFunction.SumXMY2(arr1, arr2)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 SumXMY2 方法计算两个数组中对应值之差的平方和。*/
function test() {
    Range("C1").Value2 = -5
    Range("C2").Value2 = 0
    Range("C3").Value2 = 5
    Range("C4").Value2 = 1
    Range("D1").Value2 = 23.3
    Range("D2").Value2 = 200.03
    Range("D3").Value2 = 9
    Range("D4").Value2 = 10
    console.log(WorksheetFunction.SumXMY2(Range("C1:C4"), Range("D1:D4")))
}
```
