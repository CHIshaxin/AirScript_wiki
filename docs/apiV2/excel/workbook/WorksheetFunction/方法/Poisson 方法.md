---
title: "Poisson 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Poisson%20%E6%96%B9%E6%B3%95.html
---
# Poisson (方法)

返回泊松分布。泊松分布通常用于预测一段时间内事件发生的次数，例如一分钟内通过收费站的轿车的数量。

## 说明

- 如果 x 不是整数，则将被截尾取整。
- 如果 x 或 mean 为非数字型，则 POISSON 返回错误值 #VALUE!。
- 如果 x < 0，函数 POISSON 返回错误值 #NUM!。
- 如果 mean ≤ 0，则 POISSON 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 事件数。 |
| Arg2 | double | 必填 | mean - 期望的数值。 |
| Arg3 | boolean | 必填 | cumulative - 一个决定所返回的概率分布形式的逻辑值。如果 cumulative 为 TRUE，则函数 POISSON 返回泊松累积分布概率，即，随机事件发生的次数在零到 x 之间（包括零和 x）；如果为 FALSE，则返回泊松概率密度函数，即，随机事件发生的次数正好为 x。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Poisson 方法分别计算事件数10和23的泊松分布，并分别赋值给D1和D2单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.Poisson(10, 24, true)
    Range("D2").Value2 = WorksheetFunction.Poisson(23, 12, false)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Poisson 方法计算泊松分布。*/
function test() {
    Range("B1").Value2 = 5
    Range("B2").Value2 = 3
    console.log(WorksheetFunction.Poisson(Range("B1").Value2, Range("B2").Value2, true))
}
```
