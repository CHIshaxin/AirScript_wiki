---
title: "ImSin (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImSin%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DkRphRk-
---

# ImSin (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的正弦值。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其正弦值的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImSin 方法计算以 x + yi 或 x + yj 文本格式表示的复数正弦值，并分别赋值给B6和B7单元格。*/
function test() {
    Range("B6").Value2 = WorksheetFunction.ImSin(WorksheetFunction.Complex(1, 1))
    Range("B7").Value2 = WorksheetFunction.ImSin("2+4j")
}
```

```javascript
/*本示例为B4单元格赋值，使用 ImSin 方法计算单元格B4中的复数的正弦值。*/
function test() {
    Range("B4").Formula = "=COMPLEX(1,5)"
    console.log(WorksheetFunction.ImSin(Range("B4").Value2))
}
```
