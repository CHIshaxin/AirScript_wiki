---
title: "EDate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/EDate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CM94IALe
---

# EDate (方法)

返回表示某个数据的序列号，该数据代表指定日期 (start_date) 之前或之后的月份数。使用函数 EDATE 可以计算与发行日处于一月中同一天的到期日的日期。

## 说明

- ET以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- 如果 start_date 不是有效日期，则 EDATE 将返回错误值 #VALUE!。
- 如果 months 不是整数，则将被截尾取整。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Start_date - 代表开始日期的日期。日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。 |
| Arg2 | any | 必填 | Months - start_date 之前或之后的月数。months 为正值将产生未来日期，为负值将产生过去日期。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值B1等单元格，使用 EDate 方法计算 start_date 的指定月数之后的序列号。*/
function test() {
    Range("B1").Value2 = "2008/1/1"
    Range("B2").Value2 = 1
    console.log(WorksheetFunction.EDate(Range("B1").Value2, Range("B2").Value2))
}
```

```javascript
/*本示例使用 EDate 方法计算 start_date 的指定月数之前的序列号，并分别赋值给D1等单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.EDate('2014-01-14', -1)
    Range("D2").Value2 = WorksheetFunction.EDate('2021-08-11', -5.4)
}
```
