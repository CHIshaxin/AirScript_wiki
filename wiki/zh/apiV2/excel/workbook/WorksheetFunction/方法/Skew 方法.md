---
title: "Skew (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Skew%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BNaC3HQv
---

# Skew (方法)

返回分布的不对称度。不对称度反映以其平均值为中心的分布的不对称程度。不对称度为正表示不对称部分的分布更趋向正值。不对称度为负表示不对称部分的分布更趋向负值。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，将会产生错误。
- 如果数据点数少于三个或样本标准偏差为零，SKEW 将返回错误值 #DIV/0!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2 ... - 要计算其不对称度的 1 到 30 个参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Skew 方法计算分布的不对称度，并分别赋值给C1和C2单元格。*/
function test() {
    let arr1 = [-1.2, -0.3, 1, 3, 4.5]
    Range("C1").Value2 = WorksheetFunction.Skew(arr1)
    Range("C2").Value2 = WorksheetFunction.Skew(2, 3, 6.3, 5.6, 4.3, 2.6)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Skew 方法计算分布的不对称度。*/
function test() {
    Range("B1").Value2 = 5.3
    Range("B2").Value2 = 5.8
    Range("B3").Value2 = 6.2
    Range("B4").Value2 = 7.3
    console.log(WorksheetFunction.Skew(Range("B1").Value2, Range("B2").Value2, Range("B3").Value2, Range("B4").Value2))
}
```
