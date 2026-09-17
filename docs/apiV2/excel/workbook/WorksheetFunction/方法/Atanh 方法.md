---
title: "Atanh 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Atanh%20%E6%96%B9%E6%B3%95.html
---
# Atanh (方法)

返回数字的反双曲正切值。数字必须介于 -1 和 1 之间（不包括 -1 和 1）。

## 说明

反双曲正切值的双曲正切值为 *Arg1*，因此 Atanh(Tanh(number)) 等于 *Arg1*。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 介于 1 和 -1 之间的任意实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为D3等单元格分别赋值，并使用Atanh方法计算数字的反双曲正切值。*/
function test() {
    Range("D3").Value2 = 0.9
    Range("E3").Value2 = 0
    Range("F3").Value2 = -0.1
    let atanh1 = Application.WorksheetFunction.Atanh(Range("D3").Value2)
    let atanh2 = Application.WorksheetFunction.Atanh(Range("E3").Value2)
    let atanh3 = Application.WorksheetFunction.Atanh(Range("F3").Value2)
    console.log(atanh1)
    console.log(atanh2)
    console.log(atanh3)
}
```

```javascript
/*本示例使用Atanh方法计算数字的反双曲正切值，并将结果赋值于A5等单元格。*/
function test() {
    Range("A5").Value2 = Application.WorksheetFunction.Atanh(0.5)
    Range("A6").Value2 = Application.WorksheetFunction.Atanh(0.35)
    Range("A7").Value2 = Application.WorksheetFunction.Atanh(-0.9)
    Range("A8").Value2 = Application.WorksheetFunction.Atanh(-0.4)
}
```
