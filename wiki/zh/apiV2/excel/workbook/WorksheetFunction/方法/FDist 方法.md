---
title: "FDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/FDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CwdcYsnJ
---

# FDist (方法)

返回 F 概率分布。可以使用此函数确定两个数据集是否存在变化程度上的不同。例如，可以分析进入高校的男生、女生的考试分数，确定女生分数的变化程度是否与男生不同。

## 说明

- 如果任一参数为非数值型，则 FDIST 将返回错误值 #VALUE!。
- 如果 x 为负数，则 FDIST 将返回错误值 #NUM!。
- 如果 degrees_freedom1 或 degrees_freedom2 不是整数，则将被截尾取整。
- 如果 degrees_freedom1 < 1 或 degrees_freedom1 ≥ 10^10，则 FDIST 将返回错误值 #NUM!。
- 如果 degrees_freedom2 < 1 或 degrees_freedom2 ≥ 10^10，则 FDIST 将返回错误值 #NUM!。
- FDIST 的计算公式为 FDIST=P( F>x )，其中 F 为呈 F 分布且自由度为 degrees_freedom1 和 degrees_freedom2 的随机变量。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 用于计算函数的值。 |
| Arg2 | double | 必填 | degrees_freedom1 - 分子的自由度。 |
| Arg3 | double | 必填 | degrees_freedom2 - 分母的自由度。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 FDist 方法计算 F 概率分布结果。*/
function test() {
    Range("A1").Value2 = 2.5
    Range("A2").Value2 = 5
    Range("A3").Value2 = 10
    console.log(WorksheetFunction.FDist(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```

```javascript
/*本示例使用 FDist 方法计算 F 概率分布结果，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.FDist(0.05, 7, 12)
    Range("B2").Value2 = WorksheetFunction.FDist(1.5, 30, 40)
}
```
