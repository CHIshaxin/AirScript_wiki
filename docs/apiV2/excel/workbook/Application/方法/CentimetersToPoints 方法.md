---
title: "CentimetersToPoints 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/CentimetersToPoints%20%E6%96%B9%E6%B3%95.html
---
# CentimetersToPoints (方法)

将计量单位从厘米转换为磅（一磅等于 0.035 厘米）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Centimeters | double | 必填 | 指定要转换为磅的厘米值。 |

## 返回值

Double

## 示例

```javascript
/*本示例将 Sheet1 的左边距设为 5 厘米。*/
function test() {
    Worksheets.Item("Sheet1").PageSetup.LeftMargin = Application.CentimetersToPoints(5)
}
```

```javascript
/*本示例演示将A1单元格的值用CentimetersToPoints方法转换后赋值A2单元格。*/
function test() {
    Range("A1").Value2 = 10
    Range("A2").Value2 = Application.CentimetersToPoints(Range("A1").Value2)
}
```
