---
title: "NormInv 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NormInv%20%E6%96%B9%E6%B3%95.html
---
# NormInv (方法)

返回指定平均值和标准偏差的正态累积分布函数的反函数。

## 说明

- 如果任一参数为非数字型，则 NORMINV 返回错误值 #VALUE!。
- 如果 probability < 0 或 probability > 1，则 NORMINV 返回错误值 #NUM!。
- 如果 standard_dev ≤ 0，则 NORMINV 返回错误值 #NUM!。
- 如果 mean = 0 且 standard_dev = 1，则 NORMINV 使用标准正态分布（请参阅函数 NORMSINV）。

如果已给定概率值，则 NORMINV 使用 NORMDIST(x, mean, standard_dev, TRUE) = probability 求解数值 x。因此，NORMINV 的精度取决于 NORMDIST 的精度。NORMINV 使用迭代搜索技术。如果搜索在 100 次迭代之后没有收敛，则该函数将返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | probability - 对应于正态分布的概率。 |
| Arg2 | double | 必填 | mean - 分布的算术平均值。 |
| Arg3 | double | 必填 | standard_dev - 分布的标准偏差。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NormInv 方法分别计算平均值56和标准偏差1.1、平均值12和标准偏差0.56的正态累积分布函数的反函数，并分别赋值给E1和E2单元格。*/
function test() {
    Range("E1").Value2 = WorksheetFunction.NormInv(0.132, 56, 1.1)
    Range("E2").Value2 = WorksheetFunction.NormInv(0.99, 12, 0.56)
}
```

```javascript
/*本示例为B2等单元格分别赋值，使用 NormInv 方法计算B3单元格中平均值和B4单元格中标准偏差的正态累积分布函数的反函数。*/
function test() {
    Range("B2").Value2 = 0.564
    Range("B3").Value2 = 100
    Range("B4").Value2 = 50
    console.log(WorksheetFunction.NormInv(Range("B2").Value2, Range("B3").Value2, Range("B4").Value2))
}
```
