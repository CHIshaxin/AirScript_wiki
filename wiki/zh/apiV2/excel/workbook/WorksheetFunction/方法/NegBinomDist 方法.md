---
title: "NegBinomDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NegBinomDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: xNYvLMxi
---

# NegBinomDist (方法)

返回负二项式分布。当成功概率为常量 probability_s 时，函数 NEGBINOMDIST 返回在到达 number_s 次成功之前，出现 number_f 次失败的概率。此函数与二项式分布相似，只是它的成功次数固定，试验总数为变量。与二项式分布类似的是，试验次数被假设为自变量。

## 说明

例如，如果要找 10 个反应敏捷的人，且已知具有这种特征的候选人的概率为 0.3。函数 NEGBINOMDIST 将计算出在找到 10 个合格候选人之前，需要对给定数目的不合格候选人进行面试的概率。

- number_f 和 number_s 都将被截尾取整。
- 如果任一参数为非数字型，则 NEGBINOMDIST 返回错误值 #VALUE!。
- 如果 probability_s < 0 或 probability > 1，则 NEGBINOMDIST 返回错误值 #NUM!。
- 如果 number_f < 0 或 number_s < 1，函数 NEGBINOMDIST 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | number_f - 失败的次数。 |
| Arg2 | double | 必填 | number_s - 成功次数阈值。 |
| Arg3 | double | 必填 | probability_s - 成功的概率。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NegBinomDist 方法计算负二项式分布，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.NegBinomDist(1, 2, 0.3)
    Range("A2").Value2 = WorksheetFunction.NegBinomDist(30, 50, 0.23)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 NegBinomDist 方法计算负二项式分布。*/
function test() {
    Range("B1").Value2 = 3
    Range("B2").Value2 = 5
    Range("B3").Value2 = 0.23
    console.log(WorksheetFunction.NegBinomDist(Range("B1").Value2, Range("B2").Value2, Range("B3").Value2))
}
```
