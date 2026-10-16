---
title: "Imaginary 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Imaginary%20%E6%96%B9%E6%B3%95.html
---
# Imaginary (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的虚系数。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其虚系数的复数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为使用 Imaginary 方法计算以 x + yi 或 x + yj 文本格式表示的复数的虚系数，并分别赋值给A2和A3单元格。*/
function test() {
    Range("A2").Value2 = WorksheetFunction.Imaginary(WorksheetFunction.Complex(11, 22))
    Range("A3").Value2 = WorksheetFunction.Imaginary("12-2j")
}
```

```javascript
/*本示例为A1单元格赋值，使用 Imaginary 方法计算以 x + yi 文本格式表示的A1单元格中复数的虚系数。*/
function test() {
    Range("A1").Formula2 = "=COMPLEX(3,4)"
    console.log(WorksheetFunction.Imaginary(Range("A1").Value2))
}
```
