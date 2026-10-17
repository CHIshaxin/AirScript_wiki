---
title: "ExponDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ExponDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DtQtEwPb
---

# ExponDist (方法)

返回指数分布。使用 EXPONDIST 可以建立事件之间的时间间隔模型，如银行自动提款机支付一次现金所花费的时间。例如，可以使用函数 EXPONDIST 来确定这一过程最长持续 1 分钟的发生概率。

## 说明

- 如果 x 或 lambda 为非数值型，则 EXPONDIST 将返回错误值 #VALUE!。
- 如果 x < 0，则 EXPONDIST 将返回错误值 #NUM!。
- 如果 lambda ≤ 0，则 EXPONDIST 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 函数的值。 |
| Arg2 | double | 必填 | lambda - 参数值。 |
| Arg3 | boolean | 必填 | Cumulative - 一个逻辑值，指示要提供的指数函数的形式。如果 cumulative 为 TRUE，则 EXPONDIST 将返回累积分布函数；如果为 FALSE，则它将返回概率密度函数。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 ExponDist 方法计算指数分布结果。*/
function test() {
    Range("A1").Value2 = 5
    Range("A2").Value2 = 6
    console.log(WorksheetFunction.ExponDist(Range("A1").Value2, Range("A2").Value2, false))
}
```

```javascript
/*本示例使用 ExponDist 方法计算指数分布结果，并分别赋值给C1等单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.ExponDist(5.5, 9.9, true)
    Range("C2").Value2 = WorksheetFunction.ExponDist(0, 6, false)
}
```
