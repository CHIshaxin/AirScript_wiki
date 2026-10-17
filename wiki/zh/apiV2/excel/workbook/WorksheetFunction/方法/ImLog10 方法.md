---
title: "ImLog10 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImLog10%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 62foekix
---

# ImLog10 (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的常用对数（以 10 为底数）。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其常用对数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImLog10 方法计算以 x + yi 或 x + yj 文本格式表示的复数的常用对数（以 10 为底数），并分别赋值给C3和C4单元格。*/
function test() {
    Range("C3").Value2 = WorksheetFunction.ImLog10(WorksheetFunction.Complex(6, 10))
    Range("C4").Value2 = WorksheetFunction.ImLog10("2+4i")
}
```

```javascript
/*本示例为A1单元格赋值，使用 ImLog10 方法计算单元格A1中复数的常用对数（以 10 为底数）。*/
function test() {
    Range("A1").Formula2 = "=COMPLEX(2,6)"
    console.log(WorksheetFunction.ImLog10(Range("A1").Value2))
}
```
