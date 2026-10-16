---
title: "Pearson 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Pearson%20%E6%96%B9%E6%B3%95.html
---
# Pearson (方法)

返回 Pearson（皮尔生）乘积矩相关系数 r，这是一个范围在 -1.0 到 1.0 之间（包括 -1.0 和 1.0 在内）的无量纲指数，反映了两个数据集合之间的线性相关程度。

## 说明

- 参数可以是数字，或者是包含数字的名称、数组常量或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果 array1 和 array2 为空或其数据点个数不同，函数 PEARSON 返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array1 - 一组自变量。 |
| Arg2 | any | 必填 | Array2 - 一组因变量。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Pearson 方法计算以数组arr1为自变量，数组arr2为因变量的Pearson（皮尔生）乘积矩相关系数 r，并赋值给C1单元格。*/
function test() {
    let arr1 = [1, 2, 3, 4, 5, 6]
    let arr2 = [10, 22, 33, 44, 55, 67]
    Range("C1").Value2 = WorksheetFunction.Pearson(arr1, arr2)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Pearson 方法计算以区域A1:A3为自变量，区域B1:B3为因变量的Pearson（皮尔生）乘积矩相关系数 r。*/
function test() {
    Range("A1").Value2 = 4
    Range("A2").Value2 = 3
    Range("A3").Value2 = 2
    Range("B1").Value2 = 10
    Range("B2").Value2 = 15
    Range("B3").Value2 = 18
    console.log(WorksheetFunction.Pearson(Range("A1:A3"), Range("B1:B3")))
}
```
