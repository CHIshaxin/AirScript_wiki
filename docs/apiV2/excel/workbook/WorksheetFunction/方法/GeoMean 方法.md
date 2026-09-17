---
title: "GeoMean 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/GeoMean%20%E6%96%B9%E6%B3%95.html
---
# GeoMean (方法)

返回正数数组或区域的几何平均值。例如，可以使用函数 GEOMEAN 计算可变复利的平均增长率。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果任一数据点 ≤ 0，则 GEOMEAN 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2, ... - 要计算其平均值的 1 到 30 个参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 GeoMean 方法计算正数数组的几何平均值，并分别赋值给C1和C3单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.GeoMean([1, 2, 3, 6, 9])
    Range("C3").Value2 = WorksheetFunction.GeoMean(1.3, 5.2, 6.3, 8)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 GeoMean 方法计算正数区域A1:A4的几何平均值。*/
function test() {
    Range("A1").Value2 = 11
    Range("A2").Value2 = 22
    Range("A3").Value2 = 45
    Range("A4").Value2 = 66
    console.log(WorksheetFunction.GeoMean(Range("A1:A4")))
}
```
