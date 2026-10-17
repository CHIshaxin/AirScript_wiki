---
title: "Prob (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Prob%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B9nbxi0y
---

# Prob (方法)

返回区域中的数值落在指定区间内的概率。如果没有给出上限 (upper_limit)，则返回区间 x_range 内的值等于下限 lower_limit 的概率。

## 说明

- 如果 prob_range 中的任意值 ≤ 0 或 > 1，函数 PROB 返回错误值 #NUM!。
- 如果 prob_range 中所有值之和 ¹ 1，函数 PROB 返回错误值 #NUM!。
- 如果省略 upper_limit，函数 PROB 返回值等于 lower_limit 时的概率。
- 如果 x_range 和 prob_range 中的数据点个数不同，函数 PROB 返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | X_range - 具有各自相应概率值的 x 数值区域。 |
| Arg2 | any | 必填 | Prob_range - 与 x_range 中的值相对应的一组概率值。 |
| Arg3 | double | 必填 | Lower_limit - 用于计算概率的数值下界。 |
| Arg4 | any | 可选 | Upper_limit - 用于计算概率的可选数值上界。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B1等单元格分别赋值，使用 Prob 方法计算区域中的数值落在指定区间内的概率，并分别赋值给D1和D2单元格。*/
function test() {
    Range("B1").Value2 = 5
    Range("B2").Value2 = 4
    Range("B3").Value2 = 3
    Range("B4").Value2 = 2
    Range("B5").Value2 = 1
    Range("C1").Value2 = 0.1
    Range("C2").Value2 = 0.2
    Range("C3").Value2 = 0.3
    Range("C4").Value2 = 0.2
    Range("C5").Value2 = 0.2
    Range("D1").Value2 = WorksheetFunction.Prob(Range("B1:B5"), Range("C1:C5"), 1)
    Range("D2").Value2 = WorksheetFunction.Prob(Range("B1:B5"), Range("C1:C5"), 1, 3)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Prob 方法计算区域中的数值落在指定区间内的概率。*/
function test() {
    Range("A1").Value2 = 1
    Range("A2").Value2 = 2
    Range("A3").Value2 = 3
    Range("A4").Value2 = 4
    Range("A5").Value2 = 5
    Range("B1").Value2 = 0.1
    Range("B2").Value2 = 0.25
    Range("B3").Value2 = 0.25
    Range("B4").Value2 = 0.1
    Range("B5").Value2 = 0.3
    console.log(WorksheetFunction.Prob(Range("A1:A5"), Range("B1:B5"), 2))
}
```
