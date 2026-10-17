---
title: "AmorDegrc (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/AmorDegrc%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cw1lKgqe
---

# AmorDegrc (方法)

返回每个结算期的折旧值。此函数为法国会计系统提供。

## 说明

如果某项资产是在结算期的中期购入的，则应考虑按比例折旧。该方法与 AmorLinc相似，不同之处在于该方法中用于计算的折旧系数取决于资产的寿命。

下表描述了 *Arg7* 中使用的值。

| Basis | 日期系统 |
| --- | --- |
| 0 或省略 | 360 天（NASD 方法） |
| 1 | 实际天数 |
| 3 | 一年 365 天 |
| 4 | 一年 360 天（欧洲方法） |

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- 此函数将返回折旧值，截止到资产寿命的最后一个期间，或直到累积折旧值大于资产原值减去残值后所得的结果。
- 折旧系数为： | 资产的寿命 (1/rate) | 折旧系数 | | -------------- | ---- | | 3 到 4 年 | 1.5 | | 5 到 6 年 | 2 | | 6 年以上 | 2.5 |
- 最后一个期间之前的那个期间的折旧率将增加到 50%，最后一个期间的折旧率将增加到 100%。
- 如果资产的寿命在 0（零）到 1、1 到 2、2 到 3 或 4 到 5 之间，则将返回错误值 #NUM!。

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

## 示例

```javascript
/*本示例计算每个结算期的折旧值，并将结果分别赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.AmorDegrc(4000, "2008/5/15", "2010/4/10", 100, 2, 0.3)
    Range("E3").Value2 = Application.WorksheetFunction.AmorDegrc(2500, "2015/7/10", "2016/4/15", 400, 1, 0.1, 3)
    Range("F3").Value2 = Application.WorksheetFunction.AmorDegrc(1000, "2022/1/15", "2023/1/17", 50, 1.5, 0.15, 4)
}
```

```javascript
/*本示例为B2等单元格分别赋值，再将其用于计算每个结算期的折旧值。*/
function test() {
    Range("B2").Value2 = 3000
    Range("B3").Value2 = "2021/10/1"
    Range("B4").Value2 = "2022/1/28"
    Range("B5").Value2 = 200
    Range("B6").Value2 = 1
    Range("B7").Value2 = 0.15
    Range("B8").Value2 = 1
    let amordegrc1 = Application.WorksheetFunction.AmorDegrc(Range("B2").Value2, Range("B3").Value2, Range("B4").Value2, Range("B5").Value2, Range("B6").Value2, Range("B7").Value2, Range("B8").Value2)
    console.log(amordegrc1)
}
```
