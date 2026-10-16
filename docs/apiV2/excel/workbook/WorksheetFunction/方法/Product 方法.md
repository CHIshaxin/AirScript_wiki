---
title: "Product 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Product%20%E6%96%B9%E6%B3%95.html
---
# Product (方法)

将所有以参数形式给出的数字相乘，并返回乘积值。

## 说明

- 如果参数为数字、逻辑值或数字的文本表达式，则这些参数将计算在内；如果参数为错误值或无法转换为数字的文本，则将产生错误。
- 如果参数为数组或引用，只有其中的数字将被计算。数组或引用中的空白单元格、逻辑值、文本或错误值将被忽略。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2, ... - 要相乘的 1 到 30 个数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Product 方法计算所有参数乘积，并分别赋值给C1和C2单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.Product(1, 2, 0, 3.2)
    Range("C2").Value2 = WorksheetFunction.Product(-7, -3.2, 6, 4)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Product 方法计算A1和A2单元格的乘积值。*/
function test() {
    Range("A1").Value2 = 3
    Range("A2").Value2 = 5
    console.log(WorksheetFunction.Product(Range("A1").Value2, Range("A2").Value2))
}
```
