---
title: "Choose (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Choose%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BcIIMkjb
---

# Choose (方法)

将 *Arg1* 作为索引使用，以返回值参数列表中的值。

## 说明

- 如果_Arg1_为 1，则 Choose 将返回 value1；如果为 2，则 Choose 将返回 value2，依此类推。
- 如果_Arg1_小于 1 或大于列表中最后一个值对应的数字，则 Choose 将生成一个错误。
- 如果_Arg1_是一个分数，则在使用前将被截尾取整为对应的最小整数。
- 如果_Arg1_为一个数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。），则在计算 Choose 的值时将计算每一个值。
- Choose 的值参数不仅可以为单个值，也可以为区域引用。 例如，下面的公式： =SUM(Choose(2,A1:A10,B1:B10,C1:C10)) 相当于： =SUM(B1:B10) 该函数基于区域 B1:B10 中的值返回值。 将先计算 Choose 函数的值，返回引用 B1:B10。然后使用 B1:B10 计算 SUM 函数的值。即 Choose 函数的结果是 SUM 函数的参数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 指定所选定的值参数。*Arg1* 必须为 1 到 29 之间的数字，或者是包含 1 到 29 之间的一个数字的公式，或是对包含 1 到 29 之间的一个数字的单元格的引用。 |
| Arg2 | any | 必填 | Arg30 1 到 29 个值参数，Choose 将基于 *Arg1* 从中选择一个值或要执行的一项操作。这些参数可以为数字、单元格引用、定义的名称、公式、函数或文本。 |

## 返回值

Variant

## 示例

```javascript
/*本示例使用Choose方法获取参数列表中相应的值，并将结果赋值于E3等单元格。*/
function test() {
    Range("E3").Value2 = Application.WorksheetFunction.Choose(3, "香蕉", "苹果", "桃子", "西瓜")
    Range("E4").Value2 = Application.WorksheetFunction.Choose(4, "手机", "电脑", "电视", "冰箱")
}
```

```javascript
/*本示例为B2等单元格分别赋值，并使用Choose方法获取参数列表中相应的数值。*/
function test() {
    Range("B2").Value2 = 35
    Range("B3").Value2 = 25
    Range("B4").Value2 = 12.2
    Range("B5").Value2 = 18.44
    Range("B6").Value2 = 30
    let Choose1 = Application.WorksheetFunction.Choose(2, Range("B2").Value2, Range("B3").Value2, Range("B4").Value2, Range("B5").Value2, Range("B6").Value2)
    let Choose2 = Application.WorksheetFunction.Choose(3, Range("B2").Value2, Range("B3").Value2, Range("B4").Value2, Range("B5").Value2, Range("B6").Value2)
    console.log(Choose1)
    console.log(Choose2)
}
```
