---
title: "FTest 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/FTest%20%E6%96%B9%E6%B3%95.html
---
# FTest (方法)

返回 F 检验的结果。F 检验返回的是当 array1 和 array2 的方差无明显不同时的双尾概率。可以使用此函数来确定两个样本的方差是否不同。例如，给定公立和私立学校的测试成绩，可以检验各学校间测试成绩的差别程度。

## 说明

- 参数必须是数字，或是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但包含零值的单元格将被计算在内。
- 如果 array1 或 array2 中数据点的个数小于 2，或者 array1 或 array2 的方差为零，则 FTEST 将返回错误值 #DIV/0!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | array1 - 第一个数据数组或数据区域。 |
| Arg2 | any | 必填 | array2 - 第二个数据数组或数据区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 FTest 方法根据现有的数组arr1，arr2和arr3计算 F 检验的结果，并分别赋值给D1等单元格。*/
function test() {
    let arr1 = [123, 234, 345, 456]
    let arr2 = [3, 6, 9, 15]
    let arr3 = [3, 6, 9, 15]
    Range("D1").Value2 = WorksheetFunction.FTest(arr1, arr2)
    Range("D2").Value2 = WorksheetFunction.FTest(arr2, arr3)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 FTest 方法根据现有的区域A1:A2和B1:B2中的数据计算 F 检验的结果。*/
function test() {
    Range("A1").Value2 = 12
    Range("A2").Value2 = 26
    Range("B1").Value2 = 66
    Range("B2").Value2 = 77
    console.log(WorksheetFunction.FTest(Range("A1:A2"), Range("B1:B2")))
}
```
