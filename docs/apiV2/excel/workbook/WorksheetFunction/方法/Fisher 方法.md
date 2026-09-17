---
title: "Fisher 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Fisher%20%E6%96%B9%E6%B3%95.html
---
# Fisher (方法)

返回点 x 的 Fisher 变换。该变换生成一个正态分布而非偏斜的函数。使用此函数可以完成相关系数的假设检验。

## 说明

- 如果 x 为非数值型，则 FISHER 将返回错误值 #VALUE!。
- 如果 x ≤ -1 或 x ≥ 1，则 FISHER 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 要对其进行变换的数值。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Fisher 方法计算 Fisher 转换，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Fisher(-0.9)
    Range("B2").Value2 = WorksheetFunction.Fisher(0.5)
}
```

```javascript
/*本示例分别赋值A1等单元格，使用 Fisher 方法计算A1等单元格的 Fisher 转换。*/
function test() {
    Range("A1").Value2 = 0.99
    Range("A2").Value2 = -0.9
    console.log(WorksheetFunction.Fisher(Range("A1").Value2))
    console.log(WorksheetFunction.Fisher(Range("A2").Value2))
}
```
