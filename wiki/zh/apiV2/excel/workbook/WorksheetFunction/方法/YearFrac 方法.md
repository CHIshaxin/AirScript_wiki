---
title: "YearFrac (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/YearFrac%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: ERe9Weto
---

# YearFrac (方法)

返回 start_date 和 end_date 之间的天数占全年天数的百分比。使用 YEARFRAC 工作表函数可判别某一特定条件下全年效益或债务的比例。

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
- 所有参数都将被截尾取整。
- 如果 start_date 或 end_date 不是有效日期，则 YEARFRAC 将返回错误值 #VALUE!。
- 如果 basis < 0 或 basis > 4，则 YEARFRAC 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Start_date - 代表开始日期的日期。 |
| Arg2 | any | 必填 | End_date - 代表结束日期的日期。 |
| Arg3 | any | 可选 | Basis - 要使用的日计数基准类型。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示用 YEARFRAC 计算2022/03/10和2022/11/10之间的天数占全年天数的百分比，日计数基准类型分别为默认值和3*/
function test() {
    Range("A1").Value2 = "2022/03/10"
    Range("B1").Value2 = "2022/11/10"
    console.log(WorksheetFunction.YearFrac(Range("A1").Value2, Range("B1").Value2))
    console.log(WorksheetFunction.YearFrac(Range("A1").Value2, Range("B1").Value2, 3))
}
```

```javascript
/*本示例演示用 YEARFRAC 计算2020/01/10和2022/11/10之间的天数占全年天数的百分比。*/
function test() {
    Range("A2").Value2 = WorksheetFunction.YearFrac("2020/01/10", "2022/11/10", 3)
}
```
