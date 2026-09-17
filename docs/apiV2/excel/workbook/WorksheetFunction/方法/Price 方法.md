---
title: "Price 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Price%20%E6%96%B9%E6%B3%95.html
---
# Price (方法)

返回定期付息的面值 ￥100 的债券价格。

## 说明

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

| Basis | 日计数基准 |
| --- | --- |
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1 | 实际天数/实际天数 |
| 2 | 实际天数/360 |
| 3 | 实际天数/365 |
| 4 | 欧洲 30/360 |

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- 成交日是购买者买入息票（如债券）的日期。到期日是息票有效期截止时的日期。例如，在 2008 年 1 月 1 日发行的 30 年期债券，六个月后被购买者买走。则发行日为 2008 年 1 月 1 日，成交日为 2008 年 7 月 1 日，而到期日则是在发行日 2008 年 1 月 1 日的 30 年后，即 2038 年 1 月 1 日。
- Settlement、maturity、frequency 和 basis 将被截尾取整。
- 如果 settlement 或 maturity 不是合法日期，函数 PRICE 返回错误值 #NUM!。
- 如果 yld < 0 或 rate < 0，函数 PRICE 返回错误值 #NUM!。
- 如果 redemption ≤ 0，函数 PRICE 返回错误值 #NUM!。
- 如果 frequency 不为 1、2 或 4，函数 PRICE 返回错误值 #NUM!。
- 如果 basis < 0 或 basis > 4，函数 PRICE 返回错误值 #NUM!。
- 如果 settlement ≥ maturity，函数 PRICE 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Settlement - 债券的结算日。债券的结算日是在发行日之后债券卖给购买者的日期。 |
| Arg2 | any | 必填 | Maturity - 债券的到期日。到期日是债券有效期截止时的日期。 |
| Arg3 | any | 必填 | Rate - 债券的年息票利率。 |
| Arg4 | any | 必填 | Yld - 债券的年收益。 |
| Arg5 | any | 必填 | Redemption - 面值 ￥100 的债券的赎回值。 |
| Arg6 | any | 必填 | Frequency - 每年支付息票的次数。如果按年支付，frequency = 1；如果按半年期支付，frequency = 2；如果按季支付，frequency = 4。 |
| Arg7 | any | 可选 | Basis - 要使用的日计数基准类型。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Price 方法计算定期付息的面值 ￥100 的债券价格，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Price("2020/11/20", "2023/05/12", 0.0012, 120, 15000, 1)
    Range("A2").Value2 = WorksheetFunction.Price("2019/10/11", "2025/12/12", 0.0025, 300, 30000, 2, 3)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Price 方法计算定期付息的面值 ￥100 的债券价格。*/
function test() {
    Range("A1").Value2 = "2023/04/03"
    Range("B1").Value2 = "2023/12/03"
    Range("C1").Value2 = 0.0016
    Range("D1").Value2 = 32
    Range("E1").Value2 = 10000
    Range("F1").Value2 = 4
    console.log(WorksheetFunction.Price(Range("A1").Value2, Range("B1").Value2, Range("C1").Value2, Range("D1").Value2, Range("E1").Value2, Range("F1").Value2))
}
```
