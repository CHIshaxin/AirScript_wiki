---
title: "Interior 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
---
# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，该对象在条件格式规则的计算结果为 **True** 时指定单元格的内部属性。只读。

## 示例

```javascript
/*本示例设置活动工作表上区域 C1:C10 中第二个（AboveAverage）条件格式内部的颜色。*/
function test() {
    let aboveAverage = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    aboveAverage.Interior.ColorIndex = 7
}
```

```javascript
/*本示例设置活动工作表上区域 C1:C10 中第一个（AboveAverage）条件格式内部的主题颜色。*/
function test() {
    let aboveAverage = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    aboveAverage.Interior.ThemeColor = xlThemeColorLight2
}
```
