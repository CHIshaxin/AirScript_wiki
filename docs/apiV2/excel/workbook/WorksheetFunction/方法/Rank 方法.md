---
title: "Rank 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Rank%20%E6%96%B9%E6%B3%95.html
---
# Rank (方法)

返回一个数字在数字列表中的排位。数字的排位是其大小与列表中其他值的比值（如果列表已排过序，则数字的排位就是它当前的位置）。

## 说明

- 如果 order 为 0（零）或被省略，ET 会按照 ref 为按降序排序的列表对数字排位。
- 如果 order 不为零，ET 会按照 ref 为按升序排序的列表对数字排位。
- 函数 RANK 对重复数的排位相同。但重复数的存在将影响后续数字的排位。例如，在一列按升序排列的整数中，如果数字 10 出现两次，其排位为 5，则 11 的排位为 7（没有排位为 6 的数字）。
- 由于某些原因，用户可能使用考虑重复数字的排位定义。在前面的示例中，用户可能要将整数 10 的排位改为 5.5。这可通过将下列修正系数添加到按排位返回的值来实现。此修正系数对于按照升序计算排位（顺序 = 非零值）或按照降序计算排位（顺序 = 0 或被省略）的情况都适用。 重复数排位的修正系数 =[COUNT(ref) + 1 – RANK(number, ref, 0) – RANK(number, ref, 1)]/2。 在下面的示例中，RANK(A2,A1:A5,1) 等于 3。修正系数为 (5 + 1 – 2 – 3)/2 = 0.5，考虑重复数排位后的修订排位是 3 + 0.5 = 3.5。如果数字仅在 ref 中出现一次，由于不必为重复数调整 RANK，因此修正系数为 0。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | number - 要查找其排位的数字。 |
| Arg2 | Range | 必填 | ref - 数字列表数组或对数字列表的引用。ref 中的非数值型值将被忽略。 |
| Arg3 | any | 可选 | Order - 指定数字的排位方式的数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例为C1等单元格分别赋值，使用 Rank 方法计算数值21按降序排序和11按升序排序在区域C1：C5中数字的排名，并分别赋值给D1和D2单元格。*/
function test() {
    Range("C1").Value2 = 12
    Range("C2").Value2 = 21
    Range("C3").Value2 = 10
    Range("C4").Value2 = 45
    Range("C5").Value2 = 11
    Range("D1").Value2 = WorksheetFunction.Rank(21, Range("C1:C5"), 0)
    Range("D2").Value2 = WorksheetFunction.Rank(11, Range("C1:C5"), 4)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Rank 方法计算B3单元格数值按默认降序排序在区域B1:B5中数字的排名。*/
function test() {
    Range("B1").Value2 = 6
    Range("B2").Value2 = 1
    Range("B3").Value2 = 2
    Range("B4").Value2 = 5
    Range("B5").Value2 = 10
    console.log(WorksheetFunction.Rank(Range("B3").Value2, Range("B1:B5")))
}
```
