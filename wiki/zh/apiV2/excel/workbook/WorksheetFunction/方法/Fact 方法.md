---
title: "Fact (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Fact%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D37pPyFb
---

# Fact (方法)

返回数的阶乘。一个数的阶乘等于 1*2*3*...* 该数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要计算其阶乘的非负数。如果 number 不是整数，则将被截尾取整。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 Fact 方法计算数的阶乘。*/
function test() {
    Range("A1").Value2 = 6
    Range("A2").Value2 = 5
    console.log(WorksheetFunction.Fact(Range("A1").Value2))
    console.log(WorksheetFunction.Fact(Range("A2").Value2))
}
```

```javascript
/*本示例使用 Fact 方法计算数的阶乘结果，并分别赋值给C1等单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.Fact(0)
    Range("C2").Value2 = WorksheetFunction.Fact(3.3)
}
```
