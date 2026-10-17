---
title: "FactDouble (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/FactDouble%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: yoLirmpl
---

# FactDouble (方法)

返回数字的双倍阶乘。

## 说明

- 如果 number 为非数值型，则 FACTDOUBLE 将返回错误值 #VALUE!。
- 如果 number 为负数，则 FACTDOUBLE 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要返回其双倍阶乘的值。如果 number 不是整数，则将被截尾取整。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 FactDouble 方法计算数字的双倍阶乘。*/
function test() {
    Range("A1").Value2 = 6
    Range("A2").Value2 = 5
    console.log(WorksheetFunction.FactDouble(Range("A1").Value2))
    console.log(WorksheetFunction.FactDouble(Range("A2").Value2))
}
```

```javascript
/*本示例使用 FactDouble 方法计算数字的双倍阶乘结果，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.FactDouble(0)
    Range("B2").Value2 = WorksheetFunction.FactDouble(6.6)
}
```
