---
title: "GammaDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/GammaDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DfZfz_6y
---

# GammaDist (方法)

返回 γ 分布函数值。可以使用此函数来研究具有偏态分布的变量。γ 分布通常用于队列分析。

## 说明

- 如果 x、alpha 或 beta 为非数值型，则 GAMMADIST 将返回错误值 #VALUE!。
- 如果 x < 0，则 GAMMADIST 将返回错误值 #NUM!。
- 如果 alpha ≤ 0 或 beta ≤ 0，则 GAMMADIST 将返回错误值 #NUM!。
- 对于正整数 n，当 alpha = n/2、beta = 2 且 cumulative = TRUE 时，GAMMADIST 将返回 (1 - CHIDIST(x))（自由度为 n）。
- 如果 alpha 为正整数，则 GAMMADIST 也称为 Erlang 分布。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 用于计算分布的值。 |
| Arg2 | double | 必填 | alpha - 分布参数。 |
| Arg3 | double | 必填 | beta - 分布参数。如果 beta = 1，则 GAMMADIST 将返回标准 γ 分布。 |
| Arg4 | boolean | 必填 | cumulative - 一个决定函数形式的逻辑值。如果 cumulative 为 TRUE，则 GAMMADIST 将返回累积分布函数；如果为 FALSE，则将返回概率密度函数。 |

## 返回值

Double

## 示例

```javascript
/*本示例计算γ分布函数值，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.GammaDist(1, 1, 3, true)
    Range("B2").Value2 = WorksheetFunction.GammaDist(2, 1.5, 3, false)
    Range("B3").Value2 = WorksheetFunction.GammaDist(0, 2, 3, true)
}
```

```javascript
/*本示例为A1等单元格分别赋值，并显示γ分布函数值。*/
function test() {
    Range("A1").Value2 = 2.1
    Range("A2").Value2 = 1.22
    Range("A3").Value2 = 2
    console.log(WorksheetFunction.GammaDist(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2, true))
}
```
