---
title: "Ln 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Ln%20%E6%96%B9%E6%B3%95.html
---
# Ln (方法)

返回数字的自然对数。自然对数以常数项 e (2.71828182845904) 为底。

## 说明

LN 函数是 EXP 函数的反函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要计算其自然对数的正实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Ln 方法分别计算数值99和数值154的自然对数，并分别赋值给E1和E2单元格。*/
function test() {
    Range("E1").Value2 = WorksheetFunction.Ln(99)
    Range("E2").Value2 = WorksheetFunction.Ln(154)
}
```

```javascript
/*本示例为D3单元格赋值，使用 Ln 方法计算D3单元格的自然对数。*/
function test() {
    Range("D3").Value2 = 111
    console.log(WorksheetFunction.Ln(Range("D3").Value2))
}
```
