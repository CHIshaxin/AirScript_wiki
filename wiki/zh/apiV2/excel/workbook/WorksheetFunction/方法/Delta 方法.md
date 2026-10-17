---
title: "Delta (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Delta%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DomTi4Vr
---

# Delta (方法)

检验两个值是否相等。如果 number1 = number2，则返回 1；否则返回 0。

## 说明

使用此函数可对一组值进行筛选。例如，通过对几个 DELTA 函数求和，可以计算相等的数值对的个数。此函数也称为 Kronecker Delta 函数。

- 如果 number1 为非数值型，则 DELTA 将返回错误值 #VALUE!。
- 如果 number2 为非数值型，则 DELTA 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number1 - 第一个数。 |
| Arg2 | any | 可选 | Number2 - 第二个数。如果省略，则假定 number2 为零。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A3等单元格分别赋值，并使用Delta方法检验值是否相等。*/
function test() {
    Range("A3").Value2 = 60
    Range("A4").Value2 = 30
    Range("A5").Value2 = 10
    Range("A6").Value2 = 10
    let delta1 = Application.WorksheetFunction.Delta(Range("A3").Value2, Range("A4").Value2)
    let delta2 = Application.WorksheetFunction.Delta(Range("A4").Value2)
    let delta3 = Application.WorksheetFunction.Delta(Range("A5").Value2, Range("A6").Value2)
    console.log(delta1)
    console.log(delta2)
    console.log(delta3)
}
```

```javascript
/*本示例使用Delta方法检验值是否相等，并将结果赋值于B3等单元格。*/
function test() {
    Range("B3").Value2 = Application.WorksheetFunction.Delta(10, 100)
    Range("B4").Value2 = Application.WorksheetFunction.Delta(0, 0)
    Range("B5").Value2 = Application.WorksheetFunction.Delta(12.5, 12.5)
}
```
