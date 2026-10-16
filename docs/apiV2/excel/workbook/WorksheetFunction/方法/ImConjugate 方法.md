---
title: "ImConjugate 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImConjugate%20%E6%96%B9%E6%B3%95.html
---
# ImConjugate (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的共轭复数。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其共轭复数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例为使用 ImConjugate 方法计算以 x + yi 或 x + yj 文本格式表示的复数的共轭复数，并分别赋值给C1和C2单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.ImConjugate(WorksheetFunction.Complex(1, 2))
    Range("C2").Value2 = WorksheetFunction.ImConjugate("2-4j")
}
```

```javascript
/*本示例为B1单元格赋值，使用 ImConjugate 方法计算以 x + yi 文本格式表示的复数的共轭复数。*/
function test() {
    Range("B1").Formula2 = "=COMPLEX(5,4)"
    console.log(WorksheetFunction.ImConjugate(Range("B1").Value2))
}
```
