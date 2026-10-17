---
title: "Index (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Index%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DoLWiEO6
---

# Index (方法)

返回表/区域中的值或对值的引用。INDEX 函数有两种形式：数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）形式和引用形式。

## 说明

# 数组形式

返回表元素或数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）元素的值，该元素是通过行号和列号索引选定的。

当 INDEX 的第一个参数为数组常量时，使用数组形式。

- 如果既使用 row_num 参数又使用 column_num 参数，INDEX 将返回 row_num 和 column_num 交叉处单元格的值。
- 如果将 row_num 或 column_num 设置为 0（零），INDEX 将分别返回整行或整列的值的数组。要使用以数组形式返回的值，请以数组公式 （数组公式：数组公式对一组或多组值执行多重计算，并返回一个或多个结果。数组公式括于大括号 ({ }) 中。按 Ctrl+Shift+Enter 可以输入数组公式。）形式在行的水平单元格区域或列的垂直单元格区域中输入 INDEX 函数。要输入数组公式，请按 Ctrl+Shift+Enter。
- Row_num 和 column_num 必须指向 array 中的某一单元格；否则，INDEX 将返回错误值 #REF!。

# 引用形式

返回特定的行列交叉处单元格的引用。如果该引用由非相邻的选定区域组成，则可以选择要查看的选定区域。 如果 reference 中的每个区域只包含一行或一列，则参数 row_num 或 column_num 分别为可选项。例如，对于一行的引用，可以使用 INDEX(reference,,column_num)。

- 在 reference 和 area_num 选择了特定的区域后，row_num 和 column_num 将选择一个特定的单元格：row_num 1 是该区域的第一行，column_num 1 是第一列，依此类推。INDEX 返回的引用是 row_num 和 column_num 的交叉点。
- 如果将 row_num 或 column_num 设置为 0（零），INDEX 将分别返回对整列或整行的引用。
- Row_num、column_num 和 area_num 必须指向 reference 中的某个单元格；否则，INDEX 将返回错误值 #REF!。如果省略 row_num 和 column_num，INDEX 将返回由 area_num 指定的引用中的区域。
- INDEX 函数的结果为一个引用，且在其他公式中也被解释为引用。INDEX 的返回值是作为引用还是作为值来使用要取决于公式。例如，公式 CELL("width",INDEX(A1:B2,1,2)) 与公式 CELL("width",B1) 等效。CELL 函数将 INDEX 的返回值作为单元格引用；而公式 2*INDEX(A1:B2,1,2) 将 INDEX 的返回值转换为单元格 B1 中的数字。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array 或 Reference - 一个单元格区域或一个数组常量。对引用来说，是对一个或多个单元格区域的引用。 |
| Arg2 | double | 必填 | Row_num - 数组中某行的行号，函数从该行返回值。如果 row_num 省略，则 column_num 必选。对引用来说，row_num 为引用中某行的行号，函数从该行返回一个引用。 |
| Arg3 | any | 可选 | Column_num - 数组中某列的列号，函数从该列返回值。如果 column_num 省略，则 row_num 必选。对引用来说，column_num 为引用中某列的列号，函数从该列返回一个引用。 |
| Arg4 | any | 可选 | Area_num - 仅在返回引用时使用。选择引用中的某个区域，函数从该区域返回 row_num 和 column_num 交叉处的内容。选择或输入的第一个区域的编号为 1，第二个区域为 2，依此类推。如果省略 area_num，则 INDEX 将使用区域 1。 |

## 返回值

Variant

## 示例

```javascript
/*本示例使用 Index 方法从表arr1中返回一个值，并分别赋值给A4和A5单元格。*/
function test() {
    let arr1 = [[1, 2, 3], [2, 3, 4], [0, -8, 5]]
    Range("A4").Value2 = WorksheetFunction.Index(arr1, 1, 3)
    Range("A5").Value2 = WorksheetFunction.Index(arr1, 2, 3)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Index 方法从区域A1:B2中返回一个值。*/
function test() {
    Range("A1").Value2 = 0
    Range("A2").Value2 = 9
    Range("B1").Value2 = -4
    Range("B2").Value2 = 6
    Range("C1").Value2 = 1
    Range("D1").Value2 = 2
    let arr = WorksheetFunction.Index(Range("A1:B2"), Range("C1").Value2, Range("D1").Value2)
    console.log(arr.toString())
}
```
