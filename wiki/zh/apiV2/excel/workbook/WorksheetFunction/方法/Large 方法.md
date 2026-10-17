---
title: "Large (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Large%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BC-sbrHU
---

# Large (方法)

返回数据集中第 k 个最大值。使用此函数可以根据相对标准来选择数值。例如，可以使用函数 LARGE 得到第一名、第二名或第三名的得分。

## 说明

- 如果数组为空，函数 LARGE 返回错误值 #NUM!。
- 如果 k ≤ 0 或 k 大于数据点的个数，函数 LARGE 返回错误值 #NUM!。

如果区域中数据点的个数为 n，则函数 LARGE(array,1) 返回最大值，函数 LARGE(array,n) 返回最小值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array - 要确定第 k 个最大值的数组或数据区域。 |
| Arg2 | double | 必填 | K - 要返回的数据在数组或单元格区域中的位置（从最大值开始）。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Large 方法分别计算数组arr1中第3个和第1个最大值，并分别赋值给D2和D3单元格。*/
function test() {
    let arr1 = [1, 2, 3, 4, 5, 6]
    Range("D2").Value2 = WorksheetFunction.Large(arr1, 3)
    Range("D3").Value2 = WorksheetFunction.Large(arr1, 1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Large 方法计算区域A1:A4数据集中第2个最大值。*/
function test() {
    Range("A1").Value2 = 100
    Range("A2").Value2 = 150
    Range("A3").Value2 = 70
    Range("A4").Value2 = 50
    console.log(WorksheetFunction.Large(Range("A1:A4"), 2))
}
```
