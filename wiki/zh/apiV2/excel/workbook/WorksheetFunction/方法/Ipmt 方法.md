---
title: "Ipmt (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Ipmt%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Di03gqeF
---

# Ipmt (方法)

基于固定利率和等额分期付款方式，返回给定期数内对某项投资支付的利息。

## 说明

| Type 值 | 付款时间 |
| --- | --- |
| 0 | 期末 |
| 1 | 期初 |

- 请确保用于指定 rate 和 nper 的单位的一致性。例如，同样是四年期年利率为 12% 的贷款，如果按月支付，rate 应为 12%/12，nper 应为 4*12；如果按年支付，rate 应为 12%，nper 为 4。
- 在所有参数中，支出的款项，如银行存款，表示为负数；收入的款项，如股息支票，表示为正数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Rate - 各期利率。 |
| Arg2 | double | 必填 | Per - 要计算其利息数额的期数，必须在 1 到 nper 范围内。 |
| Arg3 | double | 必填 | Nper - 年金中的总付款期数。 |
| Arg4 | double | 必填 | Pv - 现值，或一系列未来付款当前所值的累积和。 |
| Arg5 | any | 可选 | Fv - 未来值，或在最后一次付款后希望得到的现金余额。如果省略 fv，则假定其值为 0（例如，一笔贷款的未来值为 0）。 |
| Arg6 | any | 可选 | Type - 数字 0 或 1，用以指示付款时间。如果省略 type，则假定其值为 0。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Ipmt 方法计算给定期数内对某项投资支付的利息，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Ipmt(0.0036, 9, 12, -50000)
    Range("B2").Value2 = WorksheetFunction.Ipmt(0.0013, 30, 48, 10000, -1, 1)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 Ipmt 方法计算给定期数内对某项投资支付的利息。*/
function test() {
    Range("C1").Value2 = 0.00456
    Range("C2").Value2 = 24
    Range("C3").Value2 = 82
    Range("C4").Value2 = -10020
    console.log(WorksheetFunction.Ipmt(Range("C1").Value2, Range("C2").Value2, Range("C3").Value2, Range("C4").Value2))
}
```
