---
title: "TrimMean 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/TrimMean%20%E6%96%B9%E6%B3%95.html
---
# TrimMean (方法)

返回数据集的内部平均值。函数 TRIMMEAN 先从数据集的头部和尾部除去一定百分比的数据点，然后再求平均值。当希望在分析中剔除一部分数据的计算时，可以使用此函数。

## 说明

- 如果 percent < 0 或 percent > 1，则 TRIMMEAN 将返回错误值 #NUM!。
- TRIMMEAN 将除去的数据点数目向下舍入为最接近的 2 的倍数。如果 percent = 0.1，30 个数据点的 10% 等于 3 个数据点。TRIMMEAN 将对称地在数据集的头部和尾部各除去一个值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array - 要进行剪裁并求平均值的数值数组或区域。 |
| Arg2 | double | 必填 | Percent - 计算时所要排除的数据点所占的比例，例如，如果 percent = 0.2，在 20 个数据点的集合中，就要除去 4 个数据点 (20 x 0.2)：头部除去 2 个，尾部除去 2 个。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 TrimMean 方法计算数据集的内部平均值，并分别赋值给A4和A5单元格。*/
function test() {
    let arr1 = [20.232, 30.1, 40.55]
    let arr2 = [20.8, 30, 40]
    Range("A4").Value2 = WorksheetFunction.TrimMean(arr1, 0.2)
    Range("A5").Value2 = WorksheetFunction.TrimMean(arr2, 0.1)
}
```

```javascript
/*本示例为D2等单元格分别赋值，使用 TrimMean 方法计算数据集的内部平均值。*/
function test() {
    Range("D2").Value2 = 2.3
    Range("D3").Value2 = 4.3
    Range("D4").Value2 = 2.34
    console.log(WorksheetFunction.TrimMean(Range("D2:D4"), 0.1))
}
```
