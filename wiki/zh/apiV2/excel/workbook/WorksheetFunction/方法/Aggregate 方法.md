---
title: "Aggregate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Aggregate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BDcrgcZU
---

# Aggregate (方法)

返回列表或数据库中的一个聚合。

## 说明

- 以下约束将会根据 **Function_num** 值应用于 Ref 参数 (*Arg3 - Arg 30*)。 | Function_num | Ref1 | Ref2 | Ref3、Ref4… | | ------------- | ----------------------------------------------- | ----------------------------------------------- | ----------------------------------------------- | | 1-13 | **有效类型：** 任何单元格引用并集交集定义的名称结构化引用**无效类型：** 实际数据数组 | **有效类型：** 任何单元格引用并集交集定义的名称结构化引用**无效类型：** 实际数据数组 | **有效类型：** 任何单元格引用并集交集定义的名称结构化引用**无效类型：** 实际数据数组 | | 14-17 | **有效类型：** 任何单元格引用并集交集定义的名称结构化引用实际数据数组 | **有效类型：** 任何单元格引用并集交集定义的名称结构化引用实际数据数组 | **不允许使用引用** |
- 如果需要但未提供第二个 Ref 参数，AGGREGATE 将会返回错误 #VALUE!。
- 如果有一个或多个引用为三维引用，AGGREGATE 将会返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Function_num - 一个介于 1 到 19 之间的数字，用于指定要使用的函数。 Function_num 函数 1 AVERAGE 2 COUNT 3 COUNTA 4 MAX 5 MIN 6 PRODUCT 7 STDEV.S 8 STDEV.P 9 SUM 10 VAR.S 11 VAR.P 12 MEDIAN 13 MODE.SNGL 14 LARGE 15 SMALL 16 PERCENTILE.INC 17 QUARTILE.INC 18 PERCENTILE.EXC 19 QUARTILE.EXC Arg2 必选 Double 选项 - 一个数值，用于确定要在函数的求值范围中忽略的值。 选项 行为 0 或省略 忽略嵌套的 SUBTOTAL 和 AGGREGATE 函数 1 忽略隐藏行以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 2 忽略错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 3 忽略隐藏行、错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 4 忽略空值 5 忽略隐藏行 6 忽略错误值 7 忽略隐藏行和错误值 |
| Arg2 | double | 必填 | 选项 - 一个数值，用于确定要在函数的求值范围中忽略的值。 选项 行为 0 或省略 忽略嵌套的 SUBTOTAL 和 AGGREGATE 函数 1 忽略隐藏行以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 2 忽略错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 3 忽略隐藏行、错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 4 忽略空值 5 忽略隐藏行 6 忽略错误值 7 忽略隐藏行和错误值 |
| Arg3 | Range | 必填 | Ref1 - 带有多个要对其计算聚合值的数值参数的函数的第一个数值参数。 |
| Arg4 | any | 可选 | Arg 30 - Ref2 - Ref30 - 要对其计算聚合值的 2 到 30 个数值参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，再将其用于计算和、最小值以及基于样本计算标准偏差。*/
function test() {
    Range("B3").Value2 = 10
    Range("C3").Value2 = 70
    Range("D3").Value2 = 22
    Range("E3").Value2 = 46
    let sum1 = Application.WorksheetFunction.Aggregate(9, 0, Range("B3:E3"))
    let min1 = Application.WorksheetFunction.Aggregate(5, 4, Range("B3:E3"))
    let stdevs = Application.WorksheetFunction.Aggregate(7, 6, Range("B3:E3"))
    console.log(sum1)
    console.log(min1)
    console.log(stdevs)
}
```

```javascript
/*本示例为A1等单元格分别赋值，再将其用于计算平均值、最大值以及基于样本估算方差值，并将计算的值赋值于F2等单元格。*/
function test() {
    Range("A1").Value2 = 20.11
    Range("B1").Value2 = 45.6
    Range("C1").Value2 = 32.5
    Range("D1").Value2 = 78.1
    Range("F2").Value2 = Application.WorksheetFunction.Aggregate(1, 0, Range("A1:D1"))
    Range("F3").Value2 = Application.WorksheetFunction.Aggregate(5, 4, Range("A1:D1"))
    Range("F4").Value2 = Application.WorksheetFunction.Aggregate(10, 6, Range("A1:D1"))
}
```
