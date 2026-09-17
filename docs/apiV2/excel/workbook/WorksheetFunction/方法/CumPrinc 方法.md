---
title: "CumPrinc 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CumPrinc%20%E6%96%B9%E6%B3%95.html
---
# CumPrinc (方法)

返回一笔贷款在 start_period 到 end_period 期间累积支付的本金。

## 说明

下表列出了 *Arg6* 中使用的值。

| Type | 时间 |
| --- | --- |
| 0（零） | 期末付款 |
| 1 | 期初付款 |

- 请确保用于指定 rate 和 nper 的单位的一致性。例如，同样是四年期年利率为 12% 的贷款，如果按月支付，则 rate 应为 12%/12， *Arg2* 应为 4\*12；如果按年支付，则 rate 应为 12%， *Arg2* 应为 4。
- *Arg2* 、 *Arg4* 、 *Arg5* 和 type 均会被截尾取整。
- 如果 rate ≤ 0、 *Arg2* ≤ 0 或 *Arg3* ≤ 0，则 CumPrinc 将生成一个错误。
- 如果 *Arg4* < 1、 *Arg5* < 1 或 *Arg4* > *Arg5* ，则 CumPrinc 将生成一个错误。
- 如果 *Arg6* 不为数字 0 或 1，则 CumPrinc 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |

## 返回值

Double

## 示例

```javascript
/*本示例为A3等单元格分别赋值，并使用CumPrinc方法计算一笔贷款在给定的首期到末期期间累计偿还的本金数额。*/
function test() {
    Range("A3").Value2 = 0.05
    Range("A4").Value2 = 40
    Range("A5").Value2 = 125000
    Range("A6").Value2 = 20
    Range("A7").Value2 = 28
    let cumprinc1 = Application.WorksheetFunction.CumPrinc(Range("A3").Value2, Range("A4").Value2, Range("A5").Value2, Range("A6").Value2, Range("A7").Value2, 0)
    console.log(cumprinc1)
}
```

```javascript
/*本示例使用CumPrinc方法计算一笔贷款在给定的首期到末期期间累计偿还的本金数额，并将结果赋值于B3等单元格。*/
function test() {
    Range("B3").Value2 = Application.WorksheetFunction.CumPrinc(0.009, 30, 2220000, 12, 20, 0)
    Range("B4").Value2 = Application.WorksheetFunction.CumPrinc(0.04, 10, 150000, 4, 9, 1)
    Range("B5").Value2 = Application.WorksheetFunction.CumPrinc(0.07, 20, 124000, 11, 15, 1)
}
```
