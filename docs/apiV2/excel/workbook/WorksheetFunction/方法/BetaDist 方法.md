---
title: "BetaDist 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/BetaDist%20%E6%96%B9%E6%B3%95.html
---
# BetaDist (方法)

返回 Beta 累积分布函数。

## 说明

Beta 分布通常用于研究样本中某些事物变化的百分比，例如人们一天中用来看电视的时间所占的比例。

- 如果任一参数为非数值型，则 BetaDist 将返回错误值 #VALUE!。
- 如果 alpha ≤ 0 或 beta ≤ 0，则 BetaDist 将生成一个错误值。
- 如果 x < A、x > B 或 A = B，则 BetaDist 将生成一个错误值。
- 如果省略 A 和 B 的值，则 BetaDist 将使用标准累积 Beta 分布，使 A = 0，B = 1。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 介于 A 和 B 之间的值，用来计算函数的值。 |
| Arg2 | double | 必填 | 分布参数。 |
| Arg3 | double | 必填 | 分布参数。 |
| Arg4 | any | 可选 | x 所属区间的可选下界。 |
| Arg5 | any | 可选 | x 所属区间的可选上界。 |

## 返回值

Double

## 示例

```javascript
/*本示例为C3等单元格分别赋值，并使用BetaDist方法计算Beta累积分布。*/
function test() {
    Range("C3").Value2 = 2
    Range("C4").Value2 = 8
    Range("C5").Value2 = 10
    Range("C6").Value2 = 1
    Range("C7").Value2 = 3
    let betadist1 = Application.WorksheetFunction.BetaDist(Range("C3").Value2, Range("C4").Value2, Range("C5").Value2, Range("C6").Value2, Range("C7").Value2)
    console.log(betadist1)
}
```

```javascript
/*本示例使用BetaDist方法计算Beta累积分布，并将结果赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.BetaDist(3, 7, 9, 1, 6)
    Range("D4").Value2 = Application.WorksheetFunction.BetaDist(3.5, 6, 14, 2, 5)
    Range("D5").Value2 = Application.WorksheetFunction.BetaDist(2.5, 7.5, 9.5, 1, 5)
    Range("D6").Value2 = Application.WorksheetFunction.BetaDist(0.5, 7, 9)
}
```
