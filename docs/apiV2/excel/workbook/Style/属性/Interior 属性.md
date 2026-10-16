---
title: "Interior 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
---
# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，它代表指定对象的内部。

## 示例

```javascript
/*本示例将单元格 A1 的内部设置为红色。*/
function test() {
    Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 3
}
```

```javascript
/*本示例将活动工作簿第一个样式内部设置为蓝色。*/
function test() {
    ActiveWorkbook.Styles.Item(1).Interior.Color = RGB(0, 0, 255)
}
```
