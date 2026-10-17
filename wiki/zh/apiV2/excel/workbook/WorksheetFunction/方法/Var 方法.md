---
title: "Var (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Var%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 6JfXSbLK
---

# Var (方法)

估算基于样本的方差。

## 说明

- VAR 假定其参数是总体中的一个样本。如果数据代表整个总体，则应使用 VARP 来计算方差。
- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入到参数列表中的逻辑值和数字的文本表示形式也计算在内。
- 如果参数是一个数组或引用，则只计算该数组或引用中的数字。数组或引用中的空单元格、逻辑值、文本或错误值将被忽略。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - number1,number2,... - 对应于总体样本的 1 到 30 个数字参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Var 方法估算基于样本的方差，并分别赋值到A4和A5单元格。*/
function test() {
    Range("A4").Value2 = WorksheetFunction.Var(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Range("A5").Value2 = WorksheetFunction.Var(100, 200.12, 30, 4.2, 544, 6.66)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Var 方法估算基于样本的方差。*/
function test() {
    Range("A1").Value2 = -55
    Range("A2").Value2 = 100.12
    Range("A3").Value2 = 20
    Range("A4").Value2 = 1000
    console.log(WorksheetFunction.Var(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2, Range("A4").Value2))
}
```
