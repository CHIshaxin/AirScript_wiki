---
title: "MultiNomial 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/MultiNomial%20%E6%96%B9%E6%B3%95.html
---
# MultiNomial (方法)

返回参数和的阶乘与各参数阶乘乘积的比值。

## 说明

- 如果有些参数为非数值型，函数 MULTINOMIAL 返回错误值 #VALUE!。
- 如果有任何参数小于零，函数 MULTINOMIAL 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1,number2, ... - 用于进行函数 Multinomial 运算的 1 到 29 个数值参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 MultiNomial 方法计算参数和的阶乘与各参数阶乘乘积的比值，并分别赋值给A5和A6单元格。*/
function test() {
    Range("A5").Value2 = WorksheetFunction.MultiNomial(1, 2, 3)
    Range("A6").Value2 = WorksheetFunction.MultiNomial(30.3, 50.5, 60.66)
}
```

```javascript
/*本示例为B3等单元格分别赋值，使用 MultiNomial 方法计算B3等单元格的参数和的阶乘与各参数阶乘乘积的比值。*/
function test() {
    Range("B3").Value2 = 1.1
    Range("B4").Value2 = 5.5
    console.log(WorksheetFunction.MultiNomial(Range("B3").Value2, Range("B4").Value2))
}
```
