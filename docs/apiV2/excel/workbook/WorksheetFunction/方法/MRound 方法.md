---
title: "MRound 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/MRound%20%E6%96%B9%E6%B3%95.html
---
# MRound (方法)

返回参数按指定基数舍入后的数值。

## 说明

如果数值 number 除以基数的余数大于或等于基数的一半，则函数 MROUND 向远离零的方向舍入。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要进行舍入的值。 |
| Arg2 | any | 必填 | Multiple - 用以进行舍入计算的基数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 MRound 方法计算参数按指定基数舍入后的数值，并分别赋值给B4和B5单元格。*/
function test() {
    Range("B4").Value2 = WorksheetFunction.MRound(11.3, 5)
    Range("B5").Value2 = WorksheetFunction.MRound(20.35, 8)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 MRound 方法计算参数按指定基数舍入后的数值。*/
function test() {
    Range("A1").Value2 = 9.5689
    Range("A2").Value2 = 2
    console.log(WorksheetFunction.MRound(Range("A1").Value2, Range("A2").Value2))
}
```
