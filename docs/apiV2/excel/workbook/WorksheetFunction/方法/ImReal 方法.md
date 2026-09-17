---
title: "ImReal 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImReal%20%E6%96%B9%E6%B3%95.html
---
# ImReal (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的实系数。

## 说明

使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其实系数的复数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 ImReal 方法计算以 x + yi 或 x + yj 文本格式表示的复数的实系数，并分别赋值给C1和C2单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.ImReal(WorksheetFunction.Complex(12, 6))
    Range("C2").Value2 = WorksheetFunction.ImReal("10-3j")
}
```

```javascript
/*本示例为D1单元格赋值，使用 ImReal 方法计算D1单元格的复数的实系数。*/
function test() {
    Range("D1").Formula = "=COMPLEX(3,5)"
    console.log(WorksheetFunction.ImReal(Range("D1").Value2))
}
```
