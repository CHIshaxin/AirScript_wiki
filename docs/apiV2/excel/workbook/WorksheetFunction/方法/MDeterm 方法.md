---
title: "MDeterm 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/MDeterm%20%E6%96%B9%E6%B3%95.html
---
# MDeterm (方法)

返回一个数组的矩阵行列式的值。

## 说明

- Array 可以是单元格区域，例如 A1:C3；或是一个数组常量，如{1,2,3;4,5,6;7,8,9}；或是区域或数组常量的名称。
- 在以下情况下，MDTERM 返回错误值 #VALUE!：
- Array 中单元格为空或包含文字。
- Array 的行和列的数目不相等。
- 数组容量超出了 73 列 × 73 行。
- 矩阵的行列式值是由数组中的各元素计算而来的。对一个三行、三列的数组 A1:C3，其行列式的值定义如下： MDETERM(A1:C3) 等于 A1\*(B2*C3-B3*C2) + A2\*(B3*C1-B1*C3) + A3\*(B1*C2-B2*C1)。
- 矩阵的行列式值常被用来求解多元联立方程。
- 函数 MDETERM 的精确度可达十六位有效数字，因此运算结果因位数的取舍可能导致某些微小误差。例如，奇异矩阵的行列式值可能与零存在 1E-16 的误差。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array - 行数和列数相等的数值数组。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 MDeterm 方法计算数组的矩阵行列式的值，并将结果分别赋值给C1单元格。*/
function test() {
    let arr1 = [[1.2, 2.5, 3.9], [2, -3, 4], [3, 0, 5]]
    Range("C1").Value2 = WorksheetFunction.MDeterm(arr1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 MDeterm 方法计算区域A1:B2的矩阵行列式的值。*/
function test() {
    Range("A1").Value2 = 20
    Range("A2").Value2 = 40
    Range("B1").Value2 = 200
    Range("B2").Value2 = 20.6
    console.log(WorksheetFunction.MDeterm(Range("A1:B2")))
}
```
