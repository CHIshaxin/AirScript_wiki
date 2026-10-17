---
title: "ImExp (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImExp%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DrgQXfDN
---

# ImExp (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的指数。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其指数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImExp 方法计算以 x + yi 或 x + yj 文本格式表示的复数的指数，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.ImExp(WorksheetFunction.Complex(12, 11))
    Range("A2").Value2 = WorksheetFunction.ImExp("6+9j")
}
```

```javascript
/*本示例为A5单元格赋值，使用 ImExp 方法计算以 x + yi 文本格式表示的单元格A5中复数的指数。*/
function test() {
    Range("A5").Formula2 = "=COMPLEX(1,6)"
    console.log(WorksheetFunction.ImExp(Range("A5").Value2))
}
```
