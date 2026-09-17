---
title: "Weekday 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Weekday%20%E6%96%B9%E6%B3%95.html
---
# Weekday (方法)

返回与日期对应的一周中的某一天。默认情况下，其值为 1（星期日）到 7（星期六）之间的整数。

## 说明

| Return_type | 返回的数字 |
| --- | --- |
| 1 或省略 | 数字 1（星期日）到数字 7（星期六），同 ET 早期版本。 |
| 2 | 数字 1（星期一）到 7（星期日）。 |
| 3 | 数字 0（星期一）到 6（星期日）。 |
| 11 | 数字 1（星期一）到 7（星期日）。 |
| 12 | 数字 1（星期二）到 7（星期一） |
| 13 | 数字 1（星期三）到 7（星期二） |
| 14 | 数字 1（星期四）到 7（星期三） |
| 15 | 数字 1（星期五）到 7（星期四） |
| 16 | 数字 1（星期六）到 7（星期五） |
| 17 | 数字 1（星期日）到 7（星期六） |

ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Serial_number - 一个序列号，代表尝试查找的那一天的日期。日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。 |
| Arg2 | any | 可选 | Return_type - 确定返回值类型的数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示用 Weekday 计算2023/12/22与2023/12/24分别对应的一周中的某一天。*/
function test() {
    Range("C3").Value2 = WorksheetFunction.Weekday("2023/12/22", 13)
    Range("C4").Value2 = WorksheetFunction.Weekday("2023/12/24", 2)
}
```

```javascript
/*本示例演示用 Weekday 计算2023/12/20对应的一周中的某一天。*/
function test() {
    Range("B3").Value2 = "2023/12/20"
    console.log(WorksheetFunction.Weekday(Range("B3").Value2, 17))
}
```
