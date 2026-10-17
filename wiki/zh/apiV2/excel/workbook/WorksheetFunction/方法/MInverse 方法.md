---
title: "MInverse (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/MInverse%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 3GPrQ7HE
---

# MInverse (方法)

返回数组矩阵的逆距阵。

## 说明

- 数组大小不得超过 52 列 x 52 行。否则，函数会返回错误值 #VALUE!。
- Array 可以是单元格区域，例如 A1:C3；数组常量，例如 {1,2,3;4,5,6;7,8,9}；或单元格区域和数组常量的名称。
- 如果在 Array 中单元格是空白单元格或包含文字，则函数 MINVERSE 返回错误值 #VALUE!。
- 如果 Array 的行数和列数不相等，则函数 MINVERSE 也返回错误值 #VALUE!。
- 对于返回结果为数组的公式，必须以数组公式的形式输入。
- 与求行列式的值一样，求解矩阵的逆常被用于求解多元联立方程组。矩阵和它的逆矩阵相乘为单位矩阵：对角线的值为 1，其他值为 0。
- 下面是计算二阶方阵逆的示例。假设 A1:B2 中包含以字母 a、b、c 和 d 表示的四个任意的数，则下表表示矩阵 A1:B2 的逆矩阵： | 第 A 列 | 第 B 列 | | | ----- | ------------- | ------------- | | 第一行 | d/(a*d-b*c) | b/(b*c-a*d) | | 第二行 | c/(b*c-a*d) | a/(a*d-b*c) |
- 函数 MINVERSE 的精确度可达十六位有效数字，因此运算结果因位数的取舍可能会导致小的误差。
- 对于一些不能求逆的矩阵，函数 MINVERSE 将返回错误值 #NUM!。不能求逆的矩阵的行列式值为零。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array - 行数和列数相等的数值数组。 |

## 返回值

Variant

## 示例

```javascript
/*本示例为A1等单元格分别赋值，使用 MInverse 方法计算区域A1:B2矩阵的逆距阵，并将结果分别赋值给C1等单元格。*/
function test() {
    Range("A1").Value2 = 1
    Range("A2").Value2 = 2
    Range("B1").Value2 = -3
    Range("B2").Value2 = -4
    Range("C1:D2").Value2 = WorksheetFunction.MInverse(Range("A1:B2"))
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 MInverse 方法计算区域C1:D2矩阵的逆距阵。*/
function test() {
    Range("C1").Value2 = 1
    Range("C2").Value2 = 2
    Range("D1").Value2 = 3
    Range("D2").Value2 = 4
    let arr = WorksheetFunction.MInverse(Range("C1:D2"))
    console.log(arr.toString())
}
```
