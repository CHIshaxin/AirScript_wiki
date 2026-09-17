---
title: "Percentile 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Percentile%20%E6%96%B9%E6%B3%95.html
---
# Percentile (方法)

返回区域中数值的第 k 个百分点的值。可以使用此函数来建立接受阈值。例如，可以确定得分排名在第 90 个百分点之上的检测侯选人。

## 说明

- 如果 array 为空，则 PERCENTILE 将返回错误值 #NUM!。
- 如果 k 为非数字型，函数 PERCENTILE 返回错误值 #VALUE!。
- 如果 k < 0 或 k > 1，函数 PERCENTILE 返回错误值 #NUM!。
- 如果 k 不是 1/(n - 1) 的倍数，函数 PERCENTILE 使用插值法来确定第 k 个百分点的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | array - 定义相对位置的数组或数据区域。 |
| Arg2 | double | 必填 | k - 0 到 1 之间的百分点值，包括 0 和 1。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Percentile 方法分别计算数组arr数值第0.3个和第0.56个百分点的值，并分别赋值给A1和A2单元格。*/
function test() {
    let arr = [1, 2, 3, 4, 5, 6]
    Range("A1").Value2 = WorksheetFunction.Percentile(arr, 0.3)
    Range("A2").Value2 = WorksheetFunction.Percentile(arr, 0.56)
}
```

```javascript
/*本示例为C1等单元格分别赋值，显示使用 Percentile 方法计算的C1:C5单元格区域中D1单元格值对应百分点的值。*/
function test() {
    Range("C1").Value2 = 121
    Range("C2").Value2 = 141
    Range("C3").Value2 = 231
    Range("C4").Value2 = 322
    Range("C5").Value2 = 423
    Range("D1").Value2 = 0.64
    console.log(WorksheetFunction.Percentile(Range("C1:C5"), Range("D1").Value2))
}
```
