---
title: "Days360 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Days360%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CFhu2hc8
---

# Days360 (方法)

按一年 360 天的算法（一年 12 个月，每月 30 天）返回两个日期之间的天数，它用于某些会计计算中。

## 说明

如果您的会计系统是基于一年 12 个月、每月 30 天，则可使用此函数帮助计算支付款项。

下表包含了 *Arg3* 的值。

| 方法 | 定义 |
| --- | --- |
| FALSE 或省略 | 美国 (NASD) 方法。如果起始日期是一个月的 31 号，则会等于同月的 30 号。如果终止日期是一个月的 31 号，且起始日期早于 30 号，则终止日期等于下一个月的 1 号；否则，终止日期会等于同月的 30 号。 |
| TRUE | 欧洲方法。发生在一个月的 31 号的起始日期和终止日期都将等于同月的 30 号。 |

ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg2 要知道两个日期之间天数的起止日期。如果 start_date (Arg1) 发生在 end_date (Arg2) 之后，则 DAYS360 将返回一个负数。日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。 |
| Arg3 | any | 可选 | 指定在计算中是使用美国方法还是欧洲方法的一个布尔值。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，并使用Days360方法计算两个日期之间的天数。*/
function test() {
    Range("B3").Value2 = "2021-5-15"
    Range("B4").Value2 = "2021-7-18"
    Range("B5").Value2 = "2018-1-1"
    Range("B6").Value2 = "2023-5-2"
    let days3601 = Application.WorksheetFunction.Days360(Range("B3").Value2, Range("B4").Value2)
    let days3602 = Application.WorksheetFunction.Days360(Range("B5").Value2, Range("B6").Value2)
    let days3603 = Application.WorksheetFunction.Days360(Range("B5").Value2, Range("B3").Value2, true)
    console.log(days3601)
    console.log(days3602)
    console.log(days3603)
}
```

```javascript
/*本示例使用Days360方法计算两个日期之间的天数，并将结果赋值于D2等单元格。*/
function test() {
    Range("D2").Value2 = Application.WorksheetFunction.Days360("2017-2-13", "2017-1-1", true)
    Range("D3").Value2 = Application.WorksheetFunction.Days360("2015-4-12", "2016-3-15")
    Range("D4").Value2 = Application.WorksheetFunction.Days360("2022-6-19", "2024-1-1")
}
```
