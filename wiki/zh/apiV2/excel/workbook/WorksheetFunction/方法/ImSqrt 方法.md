---
title: "ImSqrt (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImSqrt%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DD-fbqdw
---

# ImSqrt (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的平方根。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其平方根的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImSqrt 方法计算以 x + yi 或 x + yj 文本格式表示的复数平方根，并分别赋值给A6和A7单元格。*/
function test() {
    Range("A6").Value2 = WorksheetFunction.ImSqrt(WorksheetFunction.Complex(6, 4))
    Range("A7").Value2 = WorksheetFunction.ImSqrt("8+8j")
}
```

```javascript
/*本示例为D1单元格赋值，使用 ImSqrt 方法计算以 x + yi 文本格式表示的复数平方根。*/
function test() {
    Range("D1").Formula = "=COMPLEX(4,2)"
    console.log(WorksheetFunction.ImSqrt(Range("D1").Value2))
}
```
