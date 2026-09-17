---
title: "Sinh 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Sinh%20%E6%96%B9%E6%B3%95.html
---
# Sinh (方法)

返回某一数字的双曲正弦值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number 为任意实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Sinh 方法计算数值0.5和2的双曲正弦值，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Sinh(0.5)
    Range("A2").Value2 = WorksheetFunction.Sinh(2)
}
```

```javascript
/*本示例为D1单元格赋值，使用 Sinh 方法计算D1单元格的双曲正弦值。*/
function test() {
    Range("D1").Value2 = 1.5
    console.log(WorksheetFunction.Sinh(Range("D1").Value2))
}
```
