---
title: "LogInv (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/LogInv%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: lJp3I-oX
---

# LogInv (方法)

使用对数分布可分析经过对数变换的数据。

## 说明

- 如果任一参数是非数字型，则 LOGINV 返回错误值 #VALUE!。
- 如果 probability <= 0 或 probability >= 1，则 LOGINV 返回错误值 #NUM!。
- 如果 standard_dev <= 0，则 LOGINV 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | probability - 与对数分布相关的概率。 |
| Arg2 | double | 必填 | mean - ln(x) 的平均值。 |
| Arg3 | double | 必填 | standard_dev - ln(x) 的标准偏差。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示 LogInv 方法使用对数分布可分析经过对数变换的数据，并分别赋值给A2和A3单元格。*/
function test() {
    Range("A2").Value2 = WorksheetFunction.LogInv(0.263, 21.3, 2)
    Range("A3").Value2 = WorksheetFunction.LogInv(0.45, 11, 0.5)
}
```

```javascript
/*本示例为A1等单元格分别赋值，演示 LogInv 方法使用对数分布可分析经过对数变换的数据。*/
function test() {
    Range("A1").Value2 = 0.75
    Range("A2").Value2 = 100
    Range("A3").Value2 = 5
    console.log(WorksheetFunction.LogInv(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
