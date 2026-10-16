---
title: "ImDiv 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImDiv%20%E6%96%B9%E6%B3%95.html
---
# ImDiv (方法)

返回以 x + yi 或 x + yj 文本格式表示的两个复数的商。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber1 - 作为分子或被除数的复数。 |
| Arg2 | any | 必填 | Inumber2 - 作为分母或除数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImDiv 方法计算以 x + yi 或 x + yj 文本格式表示的两个复数的商，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.ImDiv(WorksheetFunction.Complex(2, 11), WorksheetFunction.Complex(6, 3))
    Range("B2").Value2 = WorksheetFunction.ImDiv("4+5j", "6+5j")
}
```

```javascript
/*本示例为B3等单元格分别赋值，使用 ImDiv 方法计算以 x + yi 文本格式表示的B3和B4单元格中两个复数的商。*/
function test() {
    Range("B3").Formula2 = "=COMPLEX(2,2)"
    Range("B4").Formula2 = "=COMPLEX(2,4)"
    console.log(WorksheetFunction.ImDiv(Range("B3").Value2, Range("B4").Value2))
}
```
