---
title: "ImProduct 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImProduct%20%E6%96%B9%E6%B3%95.html
---
# ImProduct (方法)

返回以 x + yi 或 x + yj 文本格式表示的 2 到 29 个复数的乘积。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Inumber1, inumber2,... - 要相乘的 1 到 29 个复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImProduct 方法计算以 x + yi 或 x + yj 文本格式表示的2个和3个复数的乘积，并分别赋值给B3和B4单元格。*/
function test() {
    Range("B3").Value2 = WorksheetFunction.ImProduct(WorksheetFunction.Complex(2, 6), WorksheetFunction.Complex(12, 5))
    Range("B4").Value2 = WorksheetFunction.ImProduct("8+4j", "8+9j", "9+9j")
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 ImProduct 方法计算A1等单元格中复数的乘积。*/
function test() {
    Range("A1").Formula = "=COMPLEX(3,5)"
    Range("A2").Formula = "=COMPLEX(2,5)"
    Range("A3").Formula = "=COMPLEX(1,-6)"
    console.log(WorksheetFunction.ImProduct(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
