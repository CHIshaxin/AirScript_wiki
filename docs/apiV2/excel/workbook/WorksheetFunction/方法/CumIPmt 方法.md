---
title: "CumIPmt 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CumIPmt%20%E6%96%B9%E6%B3%95.html
---
# CumIPmt (方法)

返回一笔贷款在 start_period 到 end_period 期间累积支付的利息。

## 说明

下表列出了 *Arg6* 中使用的值。

| Type | 时间 |
| --- | --- |
| 0（零） | 期末付款 |
| 1 | 期初付款 |

- 请确保用于指定 rate 和 nper 的单位的一致性。例如，同样是四年期年利率为 12% 的贷款，如果按月支付，则 rate 应为 12%/12， *Arg2* 应为 4\*12；如果按年支付，则 rate 应为 12%， *Arg2* 应为 4。
- *Arg2* 、 *Arg4* 、 *Arg5* 和 type 均会被截尾取整。
- 如果 rate ≤ 0、 *Arg2* ≤ 0 或 *Arg3* ≤ 0，则 CumIPmt 将生成一个错误。
- 如果 *Arg4* < 1、*Arg5* < 1 或 *Arg4* > *Arg5* ，则 CumIPmt 将生成一个错误。
- 如果 *Arg6* 不为数字 0 或 1，则 CumIPmt 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 利率。 |
| Arg2 | any | 必填 | 总付款期数。 |
| Arg3 | any | 必填 | 现值。 |
| Arg4 | any | 必填 | 计算中的首期。付款期数从 1 开始计数。 |
| Arg5 | any | 必填 | 计算中的末期。 |
| Arg6 | any | 必填 | 付款时间。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用CumIPmt方法计算一笔贷款在给定的首期到末期期间累计偿还的利息数额，并将结果赋值于A3等单元格。*/
function test() {
    Range("A3").Value2 = Application.WorksheetFunction.CumIPmt(0.001, 40, 2000000, 1, 1, 0)
    Range("B3").Value2 = Application.WorksheetFunction.CumIPmt(0.05, 20, 1500000, 5, 8, 1)
    Range("C3").Value2 = Application.WorksheetFunction.CumIPmt(0.007, 30, 1234000, 10, 15, 1)
}
```

```javascript
/*本示例为D3等单元格分别赋值，并使用CumIPmt方法计算一笔贷款在给定的首期到末期期间累计偿还的利息数额。*/
function test() {
    Range("D3").Value2 = 0.09
    Range("D4").Value2 = 30
    Range("D5").Value2 = 1250000
    Range("D6").Value2 = 10
    Range("D7").Value2 = 20
    let cumipmt1 = Application.WorksheetFunction.CumIPmt(Range("D3").Value2, Range("D4").Value2, Range("D5").Value2, Range("D6").Value2, Range("D7").Value2, 0)
    console.log(cumipmt1)
}
```
