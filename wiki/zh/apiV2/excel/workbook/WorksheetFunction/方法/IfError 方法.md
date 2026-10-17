---
title: "IfError (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/IfError%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DYv6mzs3
---

# IfError (方法)

如果某个公式的计算结果为错误，则返回指定的一个值；否则返回该公式的结果。可以使用 IFERROR 函数捕获和处理公式 （公式：单元格中的一系列值、单元格引用、名称或运算符的组合，可生成新的值。公式总是以等号 (=) 开始。）中的错误。

## 说明

- 如果 value 或 value_if_error 是空单元格，则 IFERROR 会将其作为空字符串 ("") 处理。
- 如果 value 是数组公式，则 IFERROR 会为值中指定区域的每个单元格返回一个结果数组。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Value - 检查是否有错误的参数。 |
| Arg2 | any | 必填 | Value_if_error - 公式的计算结果为错误时要返回的值。将计算下列错误类型：#N/A、#VALUE!、#REF!、#DIV/0!、#NUM!、#NAME? 或 #NULL!。 |

## 返回值

Variant

## 示例

```javascript
/*本示例为A1等单元格分别赋值，使用 IfError 方法检查公式结果，并将正确或错误的结果分别赋值给D1和D2单元格。*/
function test() {
    Range("A1").Formula2 = 1
    Range("A2").Formula2 = 2
    Range("A3").Formula2 = 3
    Range("B1").Value2 = "=NA()"
    Range("B2").Value2 = "=sum(A1:A3)"
    Range("D1").Value2 = WorksheetFunction.IfError(Range("B1"), 12)
    Range("D2").Value2 = WorksheetFunction.IfError(Range("B2"), 22)
}
```

```javascript
/*本示例为C1等单元格分别赋值，用 IfError 方法检查公式结果，并显示该结果*/
function test() {
    Range("C1").Value2 = 10
    Range("C2").Formula2 = "=C1+7"
    console.log(WorksheetFunction.IfError(Range("C2"), 1))
}
```
