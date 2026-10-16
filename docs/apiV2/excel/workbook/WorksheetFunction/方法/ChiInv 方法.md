---
title: "ChiInv 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ChiInv%20%E6%96%B9%E6%B3%95.html
---
# ChiInv (方法)

返回 χ2 分布的单尾概率的反函数。

## 说明

如果 probability = ChiDist(x,...)，则 ChiInv(probability,...) = x。使用此函数可对观测到的结果和所期望的结果进行比较，以判定初始假设是否有效。

- 如果两个参数中的任意一个为非数值型，则 ChiInv 将生成一个错误。
- 如果 probability < 0 或 probability > 1，则 ChiInv 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom < 1 或 degrees_freedom ≥ 10^10，则 ChiInv 将生成一个错误。

如果已给定概率值，则 ChiInv 将使用 ChiDist(x, degrees_freedom) = probability 求解值 x。因此，ChiInv 的精度取决于 ChiDist 的精度。ChiInv 使用迭代搜索技术。如果搜索在 64 次迭代之后没有收敛，则该函数将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 与 χ2 分布相关的概率。 |
| Arg2 | double | 必填 | 自由度数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用ChiInv方法计算，并将结果赋值于E2等单元格。*/
function test() {
    Range("E2").Value2 = Application.WorksheetFunction.ChiInv(0.5, 5)
    Range("E3").Value2 = Application.WorksheetFunction.ChiInv(0.00055, 30)
    Range("E4").Value2 = Application.WorksheetFunction.ChiInv(0.03, 10)
}
```

```javascript
/*本示例为B2等单元格分别赋值，并使用ChiInv方法计算并显示结果。*/
function test() {
    Range("B2").Value2 = 10
    Range("B3").Value2 = 20
    Range("B4").Value2 = 0.005
    Range("B5").Value2 = 0.1
    let chidist1 = Application.WorksheetFunction.ChiInv(Range("B4").Value2, Range("B2").Value2)
    let chidist2 = Application.WorksheetFunction.ChiInv(Range("B5").Value2, Range("B3").Value2)
    console.log(chidist1)
    console.log(chidist2)
}
```
