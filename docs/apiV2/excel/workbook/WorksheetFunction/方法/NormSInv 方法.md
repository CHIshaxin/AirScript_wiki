---
title: "NormSInv 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NormSInv%20%E6%96%B9%E6%B3%95.html
---
# NormSInv (方法)

返回标准正态累积分布函数的反函数。该分布的平均值为 0，标准偏差为 1。

## 说明

- 如果 probability 为非数字型，则 NORMSINV 返回错误值 #VALUE!。
- 如果 probability <= 0 或 probability >= 1，则 NORMSINV 将返回错误值 #NUM!。

如果已给定概率值，则 NORMSINV 将使用 NORMSDIST(z) = probability 求解数值 z。因此，NORMSINV 的精度取决于 NORMSDIST 的精度。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | probability - 对应于正态分布的概率。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NormSInv 方法计算标准正态累积分布函数的反函数，并分别赋值给C3和C4单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.NormSInv(0.72)
    Range("D2").Value2 = WorksheetFunction.NormSInv(0.055)
}
```

```javascript
/*本示例为A1单元格赋值，使用 NormSInv 方法计算标准正态累积分布函数的反函数。*/
function test() {
    Range("A1").Value2 = 0.654
    console.log(WorksheetFunction.NormSInv(Range("A1").Value2))
}
```
