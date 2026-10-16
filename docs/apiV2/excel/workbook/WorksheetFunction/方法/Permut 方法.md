---
title: "Permut 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Permut%20%E6%96%B9%E6%B3%95.html
---
# Permut (方法)

返回从给定数目的对象集合中选取的若干对象的排列数。排列为有内部顺序的对象或事件的任意集合或子集。排列与组合不同，组合的内部顺序无意义。此函数可用于彩票抽奖的概率计算。

## 说明

- 两个参数将被截尾取整。
- 如果 number 或 number_chosen 为非数值型，函数 PERMUT 返回错误值 #VALUE!。
- 如果 number ≤ 0 或 number_chosen < 0，函数 PERMUT 返回错误值 #NUM!。
- 如果 number < number_chosen，函数 PERMUT 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 表示对象个数的整数。 |
| Arg2 | double | 必填 | Number_chosen - 表示每个排列中对象个数的整数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Permut 方法分别计算从10个对象集合中选取4个对象和30个对象集合中选取2个对象的排列数，并将结果分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Permut(10, 4)
    Range("B2").Value2 = WorksheetFunction.Permut(30, 2)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Permut 方法计算从给定数目的对象集合中选取的若干对象的排列数。*/
function test() {
    Range("B1").Value2 = 30
    Range("B2").Value2 = 4
    console.log(WorksheetFunction.Permut(Range("B1").Value2, Range("B2").Value2))
}
```
