---
title: "Interior 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
---
# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，它代表指定对象的内部。

## 示例

```javascript
/*本示例将活动工作表上区域 A1:A10 中第一个条件格式内部的颜色设置为红色。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.Interior.ColorIndex = 3
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式内部的主题颜色。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.Interior.ThemeColor = xlThemeColorLight2
}
```
