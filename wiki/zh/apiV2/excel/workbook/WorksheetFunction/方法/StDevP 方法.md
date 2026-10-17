---
title: "StDevP (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/StDevP%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BFnbcGYE
---

# StDevP (方法)

基于以参数形式给出的整个样本总体计算标准偏差。标准偏差反映数值相对于平均值 (mean) 的离散程度。

## 说明

- STDEVP 假定其参数为整个总体。如果数据代表总体中的样本，则应使用 STDEV 来计算标准偏差。
- 对于较大的样本大小，STDEV 和 STDEVP 的返回值大致相等。
- 此处标准偏差的计算使用“有偏差”或“n”方法。
- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入到参数列表中的逻辑值和数字的文本表示形式也计算在内。
- 如果参数是一个数组或引用，则只计算该数组或引用中的数字。数组或引用中的空单元格、逻辑值、文本或错误值将被忽略。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - number1, number2, ... - 对应于总体的 1 到 30 个数字参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 StDevP 方法获取基于以参数形式给出的整个样本总体计算标准偏差，并分别赋值给B1和B2单元格。*/
function test() {
    let arr1 = [-1.1, 2.3, 4, 5.3, 6]
    let arr2 = [-0.9, 1.3, 3, 2.3, 4]
    Range("B1").Value2 = WorksheetFunction.StDevP(arr1, arr2)
    Range("B2").Value2 = WorksheetFunction.StDevP(1, 2, 3.1, 4, 2.6, 4)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 StDevP 方法获取基于以参数形式给出的整个样本总体计算标准偏差。*/
function test() {
    Range("C1").Value2 = 5.3
    Range("C2").Value2 = 0.2
    Range("C3").Value2 = 2.3
    console.log(WorksheetFunction.StDevP(Range("C1").Value2, Range("C2").Value2, Range("C3").Value2))
}
```
