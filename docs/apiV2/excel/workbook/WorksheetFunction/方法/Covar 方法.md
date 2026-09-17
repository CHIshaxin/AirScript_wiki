---
title: "Covar 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Covar%20%E6%96%B9%E6%B3%95.html
---
# Covar (方法)

返回协方差，即每对数据点的偏差乘积的平均数。

## 说明

使用协方差可以确定两个数据集之间的关系。例如，可利用它来检验收入与教育程度是否成正比。

- 参数必须为数字，或者是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但包含零值的单元格将被计算在内。
- 如果 *Arg1* 和 *Arg2* 所含数据点的个数不同，则 COVAR 将生成一个错误。
- 如果 *Arg1* 或 *Arg2* 为空，则 Covar 将生成一个错误值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 第一个包含整数的单元格区域。 |
| Arg2 | any | 必填 | 第二个包含整数的单元格区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A3等单元格分别赋值，并使用Covar方法计算协方差。*/
function test() {
    Range("A3").Value2 = 20
    Range("A4").Value2 = 12
    Range("A5").Value2 = 34
    Range("A6").Value2 = 28
    Range("B3").Value2 = 47
    Range("B4").Value2 = 5.5
    Range("B5").Value2 = 13
    Range("B6").Value2 = 25
    let covar1 = Application.WorksheetFunction.Covar(Range("A3:A6"), Range("B3:B6"))
    console.log(covar1)
}
```

```javascript
/*本示例为C3等单元格分别赋值，并使用Covar方法计算协方差，并将结果赋值于F3单元格。*/
function test() {
    Range("C3").Value2 = 20
    Range("C4").Value2 = 12.5
    Range("C5").Value2 = 33
    Range("C6").Value2 = 27.6
    Range("D3").Value2 = 45
    Range("D4").Value2 = 5.5
    Range("D5").Value2 = 13.34
    Range("D6").Value2 = 22.6
    Range("F3").Value2 = Application.WorksheetFunction.Covar(Range("C3:C6"), Range("D3:D6"))
}
```
