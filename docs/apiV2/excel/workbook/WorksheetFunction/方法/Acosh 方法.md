---
title: "Acosh 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Acosh%20%E6%96%B9%E6%B3%95.html
---
# Acosh (方法)

返回数字的反双曲余弦值。数字必须大于或等于 1。反双曲余弦值的双曲余弦值为 *Arg1*，因此 Acosh(Cosh(number)) 等于 *Arg1*。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 等于或大于 1 的任意实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例计算数字的反双曲余弦值，并将结果分别赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.Acosh(1)
    Range("E3").Value2 = Application.WorksheetFunction.Acosh(3)
    Range("F3").Value2 = Application.WorksheetFunction.Acosh(7)
    Range("G3").Value2 = Application.WorksheetFunction.Acosh(20)
    Range("H3").Value2 = Application.WorksheetFunction.Acosh(37)
}
```

```javascript
/*本示例为c5等单元格分别赋值，再将其用于计算反双曲余弦值。*/
function test() {
    Range("C5").Value2 = 2
    Range("D5").Value2 = 34
    Range("E5").Value2 = 56
    let acosh1 = Application.WorksheetFunction.Acosh(Range("C5").Value2)
    let acosh2 = Application.WorksheetFunction.Acosh(Range("D5").Value2)
    let acosh3 = Application.WorksheetFunction.Acosh(Range("E5").Value2)
    console.log(acosh1)
    console.log(acosh2)
    console.log(acosh3)
}
```
