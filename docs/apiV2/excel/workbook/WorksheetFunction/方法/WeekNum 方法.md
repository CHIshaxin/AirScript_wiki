---
title: "WeekNum 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/WeekNum%20%E6%96%B9%E6%B3%95.html
---
# WeekNum (方法)

返回一个数字，指示一年中的第几周。

## 说明

WEEKNUM 函数将包含 1 月 1 日的一周视为一年的第一周。但有一个欧洲标准将第一周定义为周内多数天数（四天以上）都在新的一年。这意味着，对于一月份第一周的天数在三天以下的年份，按照该欧洲标准，WEEKNUM 函数返回的周数是错误的。

| Return_type | 星期开始于 |
| --- | --- |
| 1 | 星期从星期日开始。星期内的天数从 1 到 7 记数。 |
| 2 | 星期从星期一开始。星期内的天数从 1 到 7 记数。 |

ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Serial_num - 一周中的日期。日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。 |
| Arg2 | any | 可选 | Return_type - 一个数字，确定一个星期从哪一天开始。默认值为 1。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示用 WeekNum 计算2023/01/20与2023/01/03分别对应的一年中的第几周。*/
function test() {
    Range("C2").Value2 = WorksheetFunction.WeekNum("2023/01/20", 2)
    Range("C3").Value2 = WorksheetFunction.WeekNum("2023/01/03", 1)
}
```

```javascript
/*本示例演示用 WeekNum 计算2023/02/01对应的一年中的第几周。*/
function test() {
    Range("B5").Value2 = "2023/02/01"
    console.log(WorksheetFunction.WeekNum(Range("B5").Value2, 2))
}
```
