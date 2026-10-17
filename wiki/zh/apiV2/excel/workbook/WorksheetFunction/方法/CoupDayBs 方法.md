---
title: "CoupDayBs (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CoupDayBs%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DmK-uega
---

# CoupDayBs (方法)

返回从息票期开始到结算日的天数。

## 说明

下表包含了 *Arg4* 的值列表。

| Basis | 日计数基准 |
| --- | --- |
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1 | 实际天数/实际天数 |
| 2 | 实际天数/360 |
| 3 | 实际天数/365 |
| 4 | 欧洲 30/360 |

- Kingsoft WPS表格 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- 结算日是购买者买入息票（如债券）的日期。到期日是息票有效期截止时的日期。例如，在 2008 年 1 月 1 日发行的 30 年期债券，六个月后被购买者买走。则发行日为 2008 年 1 月 1 日，结算日为 2008 年 7 月 1 日，而到期日是在发行日 2008 年 1 月 1 日的 30 年后，即 2038 年 1 月 1 日。
- 所有参数都将被截尾取整。
- 如果 settlement 或 maturity 不是有效日期，则 CoupDayBs 将生成一个错误。
- 如果 frequency 不是数字 1、2 或 4，则 CoupDayBs 将生成一个错误。
- 如果 basis < 0 或 Basis > 4，则 CoupDayBs 将生成一个错误。
- 如果 settlement ≥ maturity，则 CoupDayBs 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 债券的结算日。债券的结算日是在发行日之后债券卖给购买者的日期。 |
| Arg2 | any | 必填 | 债券的到期日。到期日是债券有效期截止时的日期。 |
| Arg3 | any | 必填 | 每年支付息票的次数。如果按年支付，frequency = 1；如果按半年期支付，frequency = 2；如果按季支付，frequency = 4。 |
| Arg4 | any | 可选 | 要使用的日计数基准类型。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用CoupDayBs方法计算息票期开始到结算日的天数，并将结果赋值于B3单元格。*/
function test() {
    Range("B3").Value2 = Application.WorksheetFunction.CoupDayBs("2015-10-20", "2017-10-10", 1)
    Range("B4").Value2 = Application.WorksheetFunction.CoupDayBs("2012-5-15", "2015-6-10", 4, 3)
    Range("B5").Value2 = Application.WorksheetFunction.CoupDayBs("2010-5-10", "2018-12-27", 4, 2)
}
```

```javascript
/*本示例为D3等单元格分别赋值，并使用CoupDayBs方法计算息票期开始到结算日的天数。*/
function test() {
    Range("D3").Value2 = "2011-1-15"
    Range("D4").Value2 = "2012-11-15"
    Range("D5").Value2 = 2
    Range("D6").Value2 = 1
    let coupdaybs1 = Application.WorksheetFunction.CoupDayBs(Range("D3").Value2, Range("D4").Value2, Range("D5").Value2, Range("D6").Value2)
    console.log(coupdaybs1)
}
```
