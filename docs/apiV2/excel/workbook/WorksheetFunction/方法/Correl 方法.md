---
title: "Correl 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Correl%20%E6%96%B9%E6%B3%95.html
---
# Correl (方法)

返回单元格区域 *Arg1* 和 *Arg2* 的相关系数。

## 说明

使用相关系数可以确定两种属性之间的关系。例如，可以研究某地的平均温度和空调使用情况之间的关系。

- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果_Arg1_和_Arg2_的数据点的个数不同，则 Correl 将生成一个错误。
- 如果_Arg1_或_Arg2_为空，或者其值的 s（标准偏差）等于零，则 Correl 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 一个值单元格区域。 |
| Arg2 | any | 必填 | 另一个值单元格区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A2等单元格分别赋值，并使用Correl方法计算两个单元格区域的相关系数。*/
function test() {
    Range("A2").Value2 = 20
    Range("A3").Value2 = 30
    Range("A4").Value2 = 40
    Range("B2").Value2 = 35.5
    Range("B3").Value2 = 25
    Range("B4").Value2 = 70.1
    let correl1 = Application.WorksheetFunction.Correl(Range("A2:A4"), Range("B2:B4"))
    console.log(correl1)
}
```

```javascript
/*本示例为D3等单元格分别赋值，并使用Correl方法计算两个单元格区域的相关系数，并将结果赋值于F3单元格。*/
function test() {
    Range("D3").Value2 = 3
    Range("D4").Value2 = 2
    Range("D5").Value2 = 4
    Range("E3").Value2 = 12
    Range("E4").Value2 = 30
    Range("E5").Value2 = 50
    Range("F3").Value2 = Application.WorksheetFunction.Correl(Range("D3:D5"), Range("E3:E5"))
}
```
