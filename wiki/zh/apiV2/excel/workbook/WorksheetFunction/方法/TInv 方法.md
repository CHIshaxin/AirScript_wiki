---
title: "TInv (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/TInv%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D1HursRP
---

# TInv (方法)

返回作为概率和自由度函数的学生 t 分布的 t 值。

## 说明

- 如果两个参数中的任意一个为非数值型，则 TINV 将返回错误值 #VALUE!。
- 如果 probability < 0 或 probability > 1，则 TINV 将返回错误值 #NUM!。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom < 1，则 TINV 将返回错误值 #NUM!。
- TINV 返回 t 值，以使 P(|X| > t) = probability，其中 X 为服从 t 分布的随机变量，且 P(|X| > t) = P(X < -t or X > t)。
- 单尾 t 值可通过用两倍概率替换概率而求得。如果概率为 0.05 而自由度为 10，则双尾值由 TINV(0.05,10) 计算得到，它返回 2.28139。而同样概率和自由度的单尾值可由 TINV(2*0.05,10) 计算得到，它返回 1.812462。 在某些表格中，概率被描述为 (1-p)。 如果已给定概率值，则 TINV 使用 TDIST(x, degrees_freedom, 2) = probability 求解数值 x。因此，TINV 的精度取决于 TDIST 的精度。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Probability - 与双尾学生 t 分布关联的概率。 |
| Arg2 | double | 必填 | degrees_freedom - 代表分布的自由度数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 TInv 方法计算作为概率和自由度函数的学生 t 分布的 t 值，并分别赋值给B2和B3单元格。*/
function test() {
    Range("B2").Value2 = WorksheetFunction.TInv(0.55, 100)
    Range("B3").Value2 = WorksheetFunction.TInv(0.2, 60.235)
}
```

```javascript
/*本示例为A2等单元格分别赋值，使用 TInv 方法计算作为概率和自由度函数的学生 t 分布的 t 值。*/
function test() {
    Range("A2").Value2 = 0.3
    Range("A3").Value2 = 4.3
    console.log(WorksheetFunction.TInv(Range("A2").Value2, Range("A3").Value2))
}
```
