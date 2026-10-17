---
title: "TDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/TDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BICq7mJL
---

# TDist (方法)

返回学生 t 分布的百分点（概率），其中数值 (x) 是 t 的计算值（将计算其百分点）。t 分布用于小样本数据集的假设检验。使用此函数可以代替 t 分布的临界值表。

## 说明

- 如果任一参数为非数值型，则 TDIST 将返回错误值 #VALUE!。
- 如果 degrees_freedom < 1，则 TDIST 将返回错误值 #NUM!。
- 参数 degrees_freedom 和 tails 将被截尾取整。
- 如果 tails 不为 1 或 2，则 TDIST 将返回错误值 #NUM!。
- 如果 x < 0，则 TDIST 将返回错误值 #NUM!。
- 如果 tails = 1，则 TDIST 的计算公式为 TDIST = P( X>x )，其中 X 为服从 t 分布的随机变量。如果 tails = 2，则 TDIST 的计算公式为 TDIST = P(|X| > x) = P(X > x or X < -x)。
- 因为不允许 x < 0，当 x < 0 时要使用 TDIST，应该注意 TDIST(-x,df,1) = 1 – TDIST(x,df,1) = P(X > -x) 和 TDIST(-x,df,2) = TDIST(x df,2) = P(|X| > x)。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 用于计算分布的值。 |
| Arg2 | double | 必填 | Degrees_freedom - 表示自由度数的整数。 |
| Arg3 | double | 必填 | Tails - 指定要返回的分布的尾数。如果 tails = 1，则 TDIST 将返回单尾分布。如果 tails = 2，则 TDIST 将返回双尾分布。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 TDist 方法计算学生 t 分布的百分点（概率），并分别赋值给D2和D3单元格。*/
function test() {
    Range("D2").Value2 = WorksheetFunction.TDist(0.55, 10, 1)
    Range("D3").Value2 = WorksheetFunction.TDist(20, 1.23, 2)
}
```

```javascript
/*本示例为E1等单元格分别赋值，使用 TDist 方法计算学生 t 分布的百分点（概率）。*/
function test() {
    Range("E1").Value2 = 2
    Range("E2").Value2 = 16
    Range("F1").Value2 = 1
    console.log(WorksheetFunction.TDist(Range("E1").Value2, Range("E2").Value2, Range("F1").Value2))
}
```
