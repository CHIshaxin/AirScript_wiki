---
title: "Ispmt 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Ispmt%20%E6%96%B9%E6%B3%95.html
---
# Ispmt (方法)

计算特定投资期内要支付的利息。提供此函数是为了与 Lotus 1-2-3 兼容。

## 说明

- 请确保用于指定 rate 和 nper 的单位的一致性。例如，同样是四年期年利率为 12% 的贷款，如果按月支付，rate 应为 12%/12，nper 应为 4\*12；如果按年支付，rate 应为 12%，nper 应为 4。
- 在所有参数中，支出的金额（如银行存款或其他取款）表示为负数；接收的金额（如股息支票和其他存款）表示为正数。
- 有关财务函数的其他信息，请参阅 PV 函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Rate - 投资的利率。 |
| Arg2 | double | 必填 | Per - 要计算其利息的期数，必须介于 1 和 nper 之间。 |
| Arg3 | double | 必填 | Nper - 投资的付款期总数。 |
| Arg4 | double | 必填 | Pv - 投资的现值。对于贷款，pv 是贷款金额。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Ispmt 方法分别计算投资期10期和24期内要支付的利息，并分别赋值给D1和D2单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.Ispmt(0.0056, 10, 12, -50000)
    Range("D2").Value2 = WorksheetFunction.Ispmt(0.0023, 24, 60, 10000)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Ispmt 方法计算特定投资期内要支付的利息。*/
function test() {
    Range("A1").Value2 = 0.0023
    Range("A2").Value2 = 8
    Range("A3").Value2 = 12
    Range("A4").Value2 = 50000
    console.log(WorksheetFunction.Ispmt(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2, Range("A4").Value2))
}
```
