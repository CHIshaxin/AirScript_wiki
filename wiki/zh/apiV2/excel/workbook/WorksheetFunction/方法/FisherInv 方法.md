---
title: "FisherInv (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/FisherInv%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CfkIZSna
---

# FisherInv (方法)

返回 Fisher 变换的反函数值。使用此变换可以分析数据区域或数组之间的相关性。如果 y = FISHER(x)，则 FISHERINV(y) = x。

## 说明

- 如果 y 为非数值型，则 FISHERINV 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | y - 要对其执行逆变换的值。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 FisherInv 方法计算 Fisher 变换的反函数值，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.FisherInv(-0.9)
    Range("B2").Value2 = WorksheetFunction.FisherInv(0.5)
}
```

```javascript
/*本示例使用 FisherInv 方法计算 Fisher 变换的反函数值。*/
function test() {
    Range("A1").Value2 = 3.5
    Range("A2").Value2 = 0
    console.log(WorksheetFunction.FisherInv(Range("A1").Value2))
    console.log(WorksheetFunction.FisherInv(Range("A2").Value2))
}
```
