---
title: "AmorLinc (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/AmorLinc%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Chh3-o9Z
---

# AmorLinc (方法)

返回每个结算期的折旧值。此函数为法国会计系统提供。

## 说明

如果某项资产是在结算期的中期购入的，则应考虑按比例折旧。

下表描述了 *Arg7* 所使用的值。

| Basis | 日期系统 |
| --- | --- |
| 0 或省略 | 360 天（NASD 方法） |
| 1 | 实际天数 |
| 3 | 一年 365 天 |
| 4 | 一年 360 天（欧洲方法） |

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 资产原值。 |
| Arg2 | any | 必填 | 资产的购买日。 |
| Arg3 | any | 必填 | 第一个结算期的结束日期。 |
| Arg4 | any | 必填 | 资产在寿命结束时的残值。 |
| Arg5 | any | 必填 | 结算期。 |
| Arg6 | any | 必填 | 折旧率。 |
| Arg7 | any | 可选 | 要使用的年基准数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为G3等单元格分别赋值，再将其用于计算每个结算期的折旧值。*/
function test() {
    Range("G3").Value2 = 5000
    Range("H3").Value2 = "2002/7/1"
    Range("I3").Value2 = "2005/4/28"
    Range("J3").Value2 = 300
    Range("K3").Value2 = 1
    Range("L3").Value2 = 0.15
    Range("M3").Value2 = 1
    let amorlinc1 = Application.WorksheetFunction.AmorLinc(Range("G3").Value2, Range("H3").Value2, Range("I3").Value2, Range("J3").Value2, Range("K3").Value2, Range("L3").Value2, Range("M3").Value2)
    console.log(amorlinc1)
}
```

```javascript
/*本示例计算每个结算期的折旧值，并将结果分别赋值于A5等单元格。*/
function test() {
    Range("A5").Value2 = Application.WorksheetFunction.AmorLinc(8000, "2007/5/15", "2012/4/10", 200, 1.5, 0.3, 4)
    Range("B5").Value2 = Application.WorksheetFunction.AmorLinc(12500, "2014/4/10", "2017/8/15", 400, 1, 0.25)
    Range("C5").Value2 = Application.WorksheetFunction.AmorLinc(1000, "2021/1/15", "2021/12/17", 50, 1.5, 0.15, 3)
}
```
