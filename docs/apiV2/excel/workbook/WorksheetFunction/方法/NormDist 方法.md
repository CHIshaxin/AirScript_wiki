---
title: "NormDist 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NormDist%20%E6%96%B9%E6%B3%95.html
---
# NormDist (方法)

返回指定平均值和标准偏差的正态分布。此函数在统计方面应用范围广泛（包括假设检验）。

## 说明

- 如果 mean 或 standard_dev 为非数字型，则 NORMDIST 返回错误值 #VALUE!。
- 如果 standard_dev ≤ 0，则 NORMDIST 返回错误值 #NUM!。
- 如果 mean = 0，standard_dev = 1，且 cumulative = TRUE，则 NORMDIST 返回标准正态分布，即 NORMSDIST。
- 如果 cumulative = TRUE，则公式为从负无穷大到公式中给定的 x 的积分。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 需要计算其分布的数值。 |
| Arg2 | double | 必填 | mean - 分布的算术平均值。 |
| Arg3 | double | 必填 | standard_dev - 分布的标准偏差。 |
| Arg4 | boolean | 必填 | cumulative - 一个决定函数形式的逻辑值。如果 cumulative 为 TRUE，则 NORMDIST 返回累积分布函数；如果为 FALSE，则返回概率密度函数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NormDist 方法分别计算平均值300和标准偏差100、平均值10和标准偏差0.1的的正态分布，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.NormDist(54, 300, 100, true)
    Range("B2").Value2 = WorksheetFunction.NormDist(0.98, 10, 10.1, false)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 NormDist 方法分别计算B2单元格中平均值和B3单元格中标准偏差的的正态分布。*/
function test() {
    Range("B1").Value2 = 100
    Range("B2").Value2 = 120
    Range("B3").Value2 = 13
    console.log(WorksheetFunction.NormDist(Range("B1").Value2, Range("B2").Value2, Range("B3").Value2, true))
}
```
