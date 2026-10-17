---
title: "Slope (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Slope%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cg4T4Xoi
---

# Slope (方法)

返回根据 known_y's 和 known_x's 中的数据点拟合的线性回归直线的斜率。斜率为直线上任意两点的垂直距离与水平距离的比值，也就是回归直线的变化率。

## 说明

- 参数要么是数字，要么是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果 known_y's 和 known_x's 为空或其数据点个数不同，则 SLOPE 将返回错误值 #N/A。
- SLOPE 和 INTERCEPT 函数中使用的基础算法与 LINEST 函数中使用的基础算法不同。这些算法间的差异可能在共线数据不确定时导致不同的结果。例如，如果 known_y's 参数的数据点个数为 0，而 known_x's参数的数据点个数为 1。
- SLOPE 和 INTERCEPT 函数将返回错误 #DIV/0!。SLOPE 和 INTERCEPT 算法设计为查找一个且仅一个答案，而在此情况下可能会有多个答案。
- LINEST 将返回值 0。LINEST 算法设计为返回共线数据的合理结果，在这种情况下，至少可以找到一个答案。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 数字型相关数据点的数组或单元格区域。 |
| Arg2 | any | 必填 | Known_x's - 独立数据点的集合。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Slope 方法计算根据数组arr1和arr2中的数据点拟合的线性回归直线的斜率，并分别赋值给C3和C4单元格。*/
function test() {
    let arr1 = [1, 2, 3.2, 5]
    let arr2 = [-1.2, 0.6, 2, 3.6]
    Range("C3").Value2 = WorksheetFunction.Slope(arr1, arr2)
    Range("C4").Value2 = WorksheetFunction.Slope(arr2, arr1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Slope 方法计算根据区域A1:A4和区域B1:B4中的数据点拟合的线性回归直线的斜率。*/
function test() {
    Range("A1").Value2 = 1
    Range("A2").Value2 = 2.3
    Range("A3").Value2 = 4.5
    Range("A4").Value2 = 10
    Range("B1").Value2 = 2
    Range("B2").Value2 = 3.5
    Range("B3").Value2 = 5.7
    Range("B4").Value2 = 12
    console.log(WorksheetFunction.Slope(Range("A1:A4"), Range("B1:B4")))
}
```
