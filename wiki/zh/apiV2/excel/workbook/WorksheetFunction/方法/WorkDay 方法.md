---
title: "WorkDay (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/WorkDay%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CipvRVHL
---

# WorkDay (方法)

返回某日期（起始日期）之前或之后相隔指定工作日的某一日期的日期值。工作日不包括周末和节假日。在计算发票到期日、预期交货时间或工作天数时，可以使用函数 WORKDAY 来扣除周末或节假日。

## 说明

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- 如果任何参数不是有效日期，则 WORKDAY 将返回错误值 #VALUE。
- 如果 start_date 加 days 产生无效日期，则 WORKDAY 将返回错误值 #NUM!。
- 如果 days 不是整数，则将被截尾取整。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Start_date - 代表开始日期的日期。 |
| Arg2 | any | 必填 | Days - Start_date 之前或之后不含周末及节假日的天数。Days 为正值将产生未来日期；为负值将产生过去日期。 |
| Arg3 | any | 可选 | Holidays - 可选的列表，包含要从工作日历中排除的一个或多个日期，如各种省/市/自治区和国家/地区的法定假日和非法定假日。该列表可以是包含日期的单元格区域，也可以是由代表日期的序列号所构成的数组常量（数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示 Workday 计算起始日期为2023/12/20和2023/12/01之后分别相隔10 天、20 天工作日的某一日期的日期值，并赋值到C4等单元格。*/
function test() {
    Range("C4").Value2 = WorksheetFunction.WorkDay("2023/12/20", 10)
    Range("C5").Value2 = WorksheetFunction.WorkDay("2023/12/01", 20)
    Range("C4").NumberFormat = "yyyy/m/d"
    Range("C5").NumberFormat = "yyyy/m/d"
}
```

```javascript
/*本示例演示用 Workday 计算 2022/12/20（起始日期）之前相隔十天工作日的某一日期的日期值。*/
function test() {
    Range("A4").Value2 = "2022/12/20"
    Range("A5").Value2 = WorksheetFunction.WorkDay(Range("A4").Value2, -10)
    Range("A5").NumberFormat = "yyyy/m/d"
    console.log(Range("A5").Text)
}
```
