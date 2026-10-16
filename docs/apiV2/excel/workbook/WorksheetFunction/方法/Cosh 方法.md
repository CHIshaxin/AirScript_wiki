---
title: "Cosh 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Cosh%20%E6%96%B9%E6%B3%95.html
---
# Cosh (方法)

返回数字的双曲余弦值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 要计算其双曲余弦值的任意实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B2等单元格分别赋值，并使用Cosh方法计算数字的双曲余弦值。*/
function test() {
    Range("B2").Value2 = 4
    Range("B3").Value2 = 12
    Range("B4").Value2 = 13.5
    let cosh1 = Application.WorksheetFunction.Cosh(Range("B2").Value2)
    let cosh2 = Application.WorksheetFunction.Cosh(Range("B3").Value2)
    let cosh3 = Application.WorksheetFunction.Cosh(Range("B4").Value2)
    console.log(cosh1)
    console.log(cosh2)
    console.log(cosh3)
}
```

```javascript
/*本示例使用Cosh方法计算数字的双曲余弦值，并将结果赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.Cosh(8)
    Range("D4").Value2 = Application.WorksheetFunction.Cosh(12.3)
    Range("D5").Value2 = Application.WorksheetFunction.Cosh(1.234)
}
```
