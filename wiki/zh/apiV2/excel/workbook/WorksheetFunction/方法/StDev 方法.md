---
title: "StDev (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/StDev%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bao4FOwP
---

# StDev (方法)

基于样本估算标准偏差。标准偏差反映数值相对于平均值 (mean) 的离散程度。

## 说明

- STDEV 假定其参数是总体的样本。如果数据代表整个总体，则应使用 STDEVP 来计算标准偏差。
- 标准偏差的计算使用“无偏差”或“n-1”方法。
- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入到参数列表中的逻辑值和数字的文本表示形式也计算在内。
- 如果参数是一个数组或引用，则只计算该数组或引用中的数字。数组或引用中的空单元格、逻辑值、文本或错误值将被忽略。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - number1,number2,... - 对应于总体样本的 1 到 30 个数字参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 StDev 方法计算基于样本估算标准偏差，并分别赋值给A1和A2单元格。*/
function test() {
    let arr1 = [-111.1, 211.3, 411, 511, 611]
    let arr2 = [-110.9, 123, 345, 245.3, 411]
    Range("A1").Value2 = WorksheetFunction.StDev(arr1, arr2)
    Range("A2").Value2 = WorksheetFunction.StDev(145.2, 247.1, 345.1, 445.22, 201.26, 44)
}
```

```javascript
/*本示例为D1等单元格分别赋值，使用 StDev 方法计算基于样本估算标准偏差。*/
function test() {
    Range("D1").Value2 = 544.13
    Range("D2").Value2 = 155.23
    Range("D3").Value2 = 445.639
    console.log(WorksheetFunction.StDev(Range("D1").Value2, Range("D2").Value2, Range("D3").Value2))
}
```
