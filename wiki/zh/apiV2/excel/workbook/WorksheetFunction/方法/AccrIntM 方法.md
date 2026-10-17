---
title: "AccrIntM (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/AccrIntM%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Csk14DyY
---

# AccrIntM (方法)

返回到期一次性付息有价证券的应计利息。

## 说明

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

下表描述了可用于 *Arg5* 的值。

| 基准 | 日计数基准 |
| --- | --- |
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1 | 实际天数/实际天数 |
| 2 | 实际天数/360 |
| 3 | 实际天数/365 |
| 4 | 欧洲 30/360 |

### 以下列表包含在使用 **ACCRINTM** 时要注意的信息

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日是序列号 1，2008 年 1 月 1 日是序列号 39,448，这是因为它距 1900 年 1 月 1 日有 39,448 天。
- Issue、maturity 和 basis 将被截尾取整。
- 如果 issue 或 maturity 是无效的日期，则 ACCRINTM 将生成一个错误。
- 如果 rate ≤ 0 或 par ≤ 0，则 ACCRINTM 将产生一个错误。
- 如果 basis < 0 或 basis > 4，则 ACCRINTM 将产生一个错误。
- 如果 issue ≥ maturity，则 ACCRINTM 将产生一个错误。
- ACCRINTM 的计算公式如下： 其中： A = 按月计算的应计天数。在计算到期付息的利息时指发行日与到期日之间的天数。 D = 年基准数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 债券的发行日。 |
| Arg2 | any | 必填 | 债券的到期日。 |
| Arg3 | any | 必填 | 债券的年息票利率。 |
| Arg4 | any | 必填 | 债券的票面价值。如果省略 par，ACCRINTM 则使用 $1000。 |
| Arg5 | any | 可选 | 要使用的日计数基准类型。 |

## 返回值

Double

## 示例

```javascript
/*本示例计算到期一次性付息有价证券的应计利息，并将结果分别赋值于A3等单元格。*/
function test() {
    Range("A3").Value2 = Application.WorksheetFunction.AccrIntM("2021/3/1", "2023/5/1", 0.1, 1000)
    Range("B3").Value2 = Application.WorksheetFunction.AccrIntM("2018/2/1", "2020/7/28", 0.15, 2000, 1)
    Range("C3").Value2 = Application.WorksheetFunction.AccrIntM("2012/1/1", "2015/5/25", 0.2, 4000, 2)
}
```

```javascript
/*本示例为A2:E2分别赋值，再将其用于计算到期一次性付息有价证券的应计利息。*/
function test() {
    Range("A2").Value2 = "2015/3/11"
    Range("B2").Value2 = "2019/7/31"
    Range("C2").Value2 = 0.13
    Range("D2").Value2 = 3000
    Range("E2").Value2 = 3
    let accrintm = Application.WorksheetFunction.AccrIntM(Range("A2").Value2, Range("B2").Value2, Range("C2").Value2, Range("D2").Value2, Range("E2").Value2)
    console.log(accrintm)
}
```
