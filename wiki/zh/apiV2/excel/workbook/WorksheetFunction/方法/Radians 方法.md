---
title: "Radians (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Radians%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BYyROCgy
---

# Radians (方法)

将度转换为弧度。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Angle - 以度表示的要转换的角度。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Radians 方法分别将150度和45度转换为弧度，并分别赋值给C3和C4单元格。*/
function test() {
    Range("C3").Value2 = WorksheetFunction.Radians(150)
    Range("C4").Value2 = WorksheetFunction.Radians(45)
}
```

```javascript
/*本示例为A1单元格赋值，使用 Radians 方法将A1单元格中的度数值转换为弧度值。*/
function test() {
    Range("A1").Value2 = 175
    console.log(WorksheetFunction.Radians(Range("A1").Value2))
}
```
