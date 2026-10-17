---
title: "Confidence (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Confidence%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CCsaI4gr
---

# Confidence (方法)

返回一个值，可使用该值构造总体平均值的置信区间。

## 说明

置信区间为一个值区域。样本平均值 x 位于此区域的中心，区域范围为 x ± Confidence。例如，如果通过邮件订购产品，其交付时间的样本平均值为 x，则总体平均值的区域范围为 x ± Confidence。对于任何包含在此区域中的总体平均值 μ0，从 μ0 到 x，获取样本平均值的概率大于 alpha；对于任何未在此区域中的总体平均值 μ0，从 μ0 到 x，获取样本平均值的概率小于 alpha。换句话说，假设使用 x、standard_dev 和 size 构造一个双尾检验，假设的显著性水平为 alpha，总体平均值为 μ0。如果 μ0 在置信区间中，则将不会否决该假设；如果 μ0 不在置信区间中，则将否决该假设。置信区间不允许进行概率为 1 - alpha 的推断，此时下一份包裹的交付时间将肯定位于置信区间内。

- 如果任一参数为非数值型，则 Confidence 将生成一个错误。
- 如果 alpha ≤ 0 或 alpha ≥ 1，则 Confidence 将生成一个错误。
- 如果 standard_dev ≤ 0，则 Confidence 将生成一个错误。
- 如果 size 不是整数，则将被截尾取整。
- 如果 size < 1，则 Confidence 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 用于计算置信度的显著性水平。置信度等于 100*(1 - alpha)%，亦即，如果 alpha 为 0.05，则置信度为 95%。 |
| Arg2 | double | 必填 | 数据区域的总体标准偏差，假定为已知。 |
| Arg3 | double | 必填 | 样本大小。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用Confidence方法计算总体平均值的置信区间，并将结果赋值于C3等单元格。*/
function test() {
    Range("C3").Value2 = Application.WorksheetFunction.Confidence(0.05, 2.5, 50)
    Range("C4").Value2 = Application.WorksheetFunction.Confidence(0.4, 3, 60)
    Range("C5").Value2 = Application.WorksheetFunction.Confidence(0.001, 1.5, 40)
}
```

```javascript
/*本示例为D2等单元格分别赋值，并使用Confidence方法计算总体平均值的置信区间。*/
function test() {
    Range("D2").Value2 = 0.1
    Range("D3").Value2 = 0.0005
    Range("D4").Value2 = 3.5
    Range("D5").Value2 = 1.4
    Range("D6").Value2 = 50
    Range("D7").Value2 = 30
    let confidence1 = Application.WorksheetFunction.Confidence(Range("D2").Value2, Range("D4").Value2, Range("D6").Value2)
    let confidence2 = Application.WorksheetFunction.Confidence(Range("D3").Value2, Range("D5").Value2, Range("D7").Value2)
    console.log(confidence1)
    console.log(confidence2)
}
```
