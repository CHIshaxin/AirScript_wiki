---
title: "ErfC 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ErfC%20%E6%96%B9%E6%B3%95.html
---
# ErfC (方法)

返回从指定参数到无穷积分的 ERF 函数的余误差函数。

## 说明

如果该函数不可用，并返回 #NAME? 错误，则需要安装并加载\*\*“分析工具库”\*\*加载宏。

如果参数是非数值，则 ERFC 返回 #VALUE! 错误值。

如果参数是负值，则 ERFC 返回 #NUM! 错误值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 第一个参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值B1等单元格，使用 ErfC 方法计算指定参数到无穷积分的 ERF 函数的余误差函数。*/
function test() {
    Range("B1").Value2 = 1
    Range("B2").Value2 = 3.3
    console.log(WorksheetFunction.ErfC(Range("B1").Value2))
    console.log(WorksheetFunction.ErfC(Range("B2").Value2))
}
```

```javascript
/*本示例使用 ErfC 方法计算指定参数到无穷积分的 ERF 函数的余误差函数，并分别赋值给D1等单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.ErfC(0)
    Range("D2").Value2 = WorksheetFunction.ErfC(1000)
}
```
