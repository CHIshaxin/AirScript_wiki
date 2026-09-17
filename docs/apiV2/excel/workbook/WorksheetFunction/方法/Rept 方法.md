---
title: "Rept 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Rept%20%E6%96%B9%E6%B3%95.html
---
# Rept (方法)

按照给定的次数重复显示文本。可以通过函数 REPT 来不断地重复显示某一文本字符串，对单元格进行填充。

## 说明

- 如果 number_times 为 0（零），则 REPT 将返回 ""（空文本）。
- 如果 number_times 不是整数，则将被截尾取整。
- REPT 函数的结果不能大于 32,767 个字符，否则 REPT 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | Text - 要重复的文本。 |
| Arg2 | double | 必填 | Number_times - 指定文本重复次数的一个正数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Rept 方法分别重复显示“加油”3次和“欢迎”5次，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Rept("加油", 3)
    Range("B2").Value2 = WorksheetFunction.Rept("欢迎", 5)
}
```

```javascript
/*本示例为C1等单元格赋值，使用 Rept 方法按照D1单元格给定的次数重复显示C1单元格中的文本。*/
function test() {
    Range("C1").Value2 = "欢迎您！"
    Range("D1").Value2 = 2
    console.log(WorksheetFunction.Rept(Range("C1").Value2, Range("D1").Value2))
}
```
