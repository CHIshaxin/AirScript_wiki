---
title: "NormSDist 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NormSDist%20%E6%96%B9%E6%B3%95.html
---
# NormSDist (方法)

返回标准正态累积分布函数，该分布的平均值为 0（零），标准偏差为 1。可以使用此函数代替标准正态曲线面积表。

## 说明

- 如果 z 为非数字型，则 NORMSDIST 返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | z - 要计算其分布的值。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NormSDist 方法计算标准正态累积分布函数，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.NormSDist(0.56)
    Range("A2").Value2 = WorksheetFunction.NormSDist(0.23)
}
```

```javascript
/*本示例为D1单元格赋值，使用 NormSDist 方法计算标准正态累积分布函数。*/
function test() {
    Range("D1").Value2 = 0.456
    console.log(WorksheetFunction.NormSDist(Range("D1").Value2))
}
```
