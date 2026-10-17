---
title: "BinomDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/BinomDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D0YGox26
---

# BinomDist (方法)

返回一元二项式分布的概率。

## 说明

当任何试验的结果仅包含成功或失败两种情况，当试验是独立试验，以及当整个试验过程中成功的概率固定不变时，BinomDist 适用于固定次数的检验或试验。例如，BinomDist 可以计算出生的三个婴儿中两个是男孩的概率。

- Number_s 和 trials 将被截尾取整。
- 如果 number_s、trials 或 probability_s 为非数值型，则 BinomDist 将生成一个错误。
- 如果 number_s < 0 或 number_s > trials，则 BinomDist 将生成一个错误。
- 如果 probability_s < 0 或 probability_s > 1，则 BinomDist 将生成一个错误。
- 二项式概率密度函数的计算公式为： 其中： 为 COMBIN(n,x)。 累积二项式分布的计算公式为：

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 试验成功的次数。 |
| Arg2 | double | 必填 | 独立试验的次数 |
| Arg3 | double | 必填 | 每次试验成功的概率。 |
| Arg4 | boolean | 必填 | 一个逻辑值，用于确定函数的形式。如果 cumulative 为 TRUE，则 BinomDist 返回累积分布函数，即最多存在 number_s 次成功的概率；如果为 FALSE，则返回概率密度函数，即存在 number_s 次成功的概率。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，并使用BinomDist方法返回一元二项式分布的概率。*/
function test() {
    Range("B3").Value2 = 17
    Range("C3").Value2 = 20
    Range("D3").Value2 = 0.7
    let binomdist1 = Application.WorksheetFunction.BinomDist(Range("B3").Value2, Range("C3").Value2, Range("D3").Value2, true)
    let binomdist2 = Application.WorksheetFunction.BinomDist(Range("B3").Value2, Range("C3").Value2, Range("D3").Value2, false)
    console.log(binomdist1)
    console.log(binomdist2)
}
```

```javascript
/*本示例使用BinomDist方法返回一元二项式分布的概率，并将结果赋值于D4等单元格。*/
function test() {
    Range("D4").Value2 = Application.WorksheetFunction.BinomDist(8, 10, 0.9, true)
    Range("D5").Value2 = Application.WorksheetFunction.BinomDist(10, 30, 0.5, false)
    Range("D6").Value2 = Application.WorksheetFunction.BinomDist(20, 20, 0.8, true)
}
```
