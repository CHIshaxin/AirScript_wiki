---
title: "ImLog2 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImLog2%20%E6%96%B9%E6%B3%95.html
---
# ImLog2 (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的以 2 为底数的对数。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其以 2 为底数的对数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImLog2 方法计算以 x + yi 或 x + yj 文本格式表示的复数以2为底数的对数，并分别赋值给B3和B4单元格。*/
function test() {
    Range("B3").Value2 = WorksheetFunction.ImLog2(WorksheetFunction.Complex(6, 9))
    Range("B4").Value2 = WorksheetFunction.ImLog2("12+5j")
}
```

```javascript
/*本示例为B1单元格赋值，使用 ImLog2 方法计算B1单元格中复数的以2为底数的对数。*/
function test() {
    Range("B1").Formula2 = "=COMPLEX(8,8)"
    console.log(WorksheetFunction.ImLog2(Range("B1").Value2))
}
```
