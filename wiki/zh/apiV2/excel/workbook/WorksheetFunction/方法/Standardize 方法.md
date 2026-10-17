---
title: "Standardize (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Standardize%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CLA9VzmS
---

# Standardize (方法)

返回以 mean 和 standard_dev 表示的分布的正态化数值。

## 说明

- 如果 standard_dev ≤ 0，则 STANDARDIZE 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | X - 要进行正态化的值。 |
| Arg2 | double | 必填 | Mean - 分布的算术平均值。 |
| Arg3 | double | 必填 | Standard_dev - 分布的标准偏差。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Standardize 方法计算以 mean 和 standard_dev 表示的分布的正态化数值，并分别赋值给A3和B3单元格。*/
function test() {
    Range("A3").Value2 = WorksheetFunction.Standardize(-2.3, 1, 1.2)
    Range("B3").Value2 = WorksheetFunction.Standardize(1, 2.2, 8.3)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Standardize 方法计算以 mean 和 standard_dev 表示的分布的正态化数值。*/
function test() {
    Range("A1").Value2 = 7.2
    Range("B1").Value2 = 2
    Range("C1").Value2 = 3.3
    console.log(WorksheetFunction.Standardize(Range("A1").Value2, Range("B1").Value2, Range("C1").Value2))
}
```
