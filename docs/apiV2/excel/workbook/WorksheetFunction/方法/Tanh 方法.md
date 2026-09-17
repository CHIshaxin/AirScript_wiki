---
title: "Tanh 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Tanh%20%E6%96%B9%E6%B3%95.html
---
# Tanh (方法)

返回某一数字的双曲正切值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 任意实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Tanh 方法计算数值1和2.5的双曲正切值，并分别赋值给B2和B3单元格。*/
function test() {
    Range("B2").Value2 = WorksheetFunction.Tanh(1)
    Range("B3").Value2 = WorksheetFunction.Tanh(2.5)
}
```

```javascript
/*本示例为A3单元格赋值，使用 Tanh 方法计算其双曲正切值。*/
function test() {
    Range("A3").Value2 = -1
    console.log(WorksheetFunction.Tanh(Range("A3").Value2))
}
```
