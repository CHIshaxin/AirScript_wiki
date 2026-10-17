---
title: "NetworkDays (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NetworkDays%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cnh-JToo
---

# NetworkDays (方法)

返回参数 start-data 和 end-data 之间的全部工作日数值。工作日不包括周末和节假日。可以使用函数 NETWORKDAYS，根据某一特定时期内雇员的工作天数，计算其应计的报酬。

## 说明

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- 如果任何参数为无效的日期值，则函数 NETWORKDAYS 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Start_date - 代表开始日期的日期。 |
| Arg2 | any | 必填 | End_date - 代表结束日期的日期。 |
| Arg3 | any | 可选 | Holidays - 可选的区域，包含要从工作日历中排除的一个或多个日期，如各种省/市/自治区和国家/地区的法定假日和非法定假日。该列表可以是包含日期的单元格区域，也可以是由代表日期的序列号所构成的数组常量 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NetworkDays 方法分别计算日期“2024/01/01”和“2024/01/10”、“2024/01/02”和“2024/01/19”之间的全部工作日数值，并分别赋值给B3和B4单元格。*/
function test() {
    Range("B3").Value2 = WorksheetFunction.NetworkDays("2024/01/01", "2024/01/10", "2024/01/01")
    Range("B4").Value2 = WorksheetFunction.NetworkDays("2024/01/02", "2024/01/19")
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 NetworkDays 方法计算A1单元格中日期和A2单元格中日期之间的全部工作日数值。*/
function test() {
    Range("A1").Value2 = "2023/12/01"
    Range("A2").Value2 = "2024/01/02"
    Range("A3").Value2 = "2024/01/01"
    console.log(WorksheetFunction.NetworkDays(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
