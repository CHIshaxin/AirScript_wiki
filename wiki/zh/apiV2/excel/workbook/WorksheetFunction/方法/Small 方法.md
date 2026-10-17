---
title: "Small (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Small%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D-ngyfow
---

# Small (方法)

返回数据集中第 k 个最小值。使用此函数可以返回数据集中特定位置上的数值。

## 说明

- 如果 array 为空，则 SMALL 将返回错误值 #NUM!。
- 如果 k ≤ 0 或 k 超过了数据点个数，则 SMALL 将返回错误值 #NUM!。
- 如果 n 为数组中数据点的个数，则 SMALL(array,1) 等于最小值，SMALL(array,n) 等于最大值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array - 需要确定第 k 个最小值的数值数据数组或区域。 |
| Arg2 | double | 必填 | K - 要返回的数据在数组或区域中的位置（从最小值开始）。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Small 方法计算数组中第1个和第3个最小值，并分别赋值给B2和B3单元格。*/
function test() {
    let arr1 = [1, 2, 3, 2.5, 5]
    Range("B2").Value2 = WorksheetFunction.Small(arr1, 1)
    Range("B3").Value2 = WorksheetFunction.Small(arr1, 3)
}
```

```javascript
/*本示例为A1等单元格赋值，使用 Small 方法计算A1:A4区域中第2个最小值。*/
function test() {
    Range("A1").Value2 = 1.6
    Range("A2").Value2 = 5.6
    Range("A3").Value2 = 3.6
    Range("A4").Value2 = 7.6
    console.log(WorksheetFunction.Small(Range("A1:A4"), 2))
}
```
