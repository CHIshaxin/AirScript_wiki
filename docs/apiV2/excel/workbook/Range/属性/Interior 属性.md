---
title: "Interior 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
---
# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，它代表指定对象的内部。

## 示例

```javascript
/*此示例将 Sheet1 中 A1 单元格的内部颜色设为蓝绿色。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").Interior.ColorIndex = 8 // Cyan
}
```

```javascript
/*本示例将 C5 单元格的颜色设置为红色，并将颜色变为最深*/
function test() {
    let interior = Range("C5").Interior
    interior.Color = RGB(255, 0, 0)
    interior.TintAndShade = -1
}
```
