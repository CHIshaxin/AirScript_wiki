---
title: "Borders 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
---
# Borders (属性)

返回一个 **[Borders](./../../Borders/Borders 对象.html)** 对象，该对象代表当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的边框。只读。

## 返回值

Borders

## 示例

```javascript
/*此示例显示 B4 单元格边框在当前调色板中颜色的索引值是否为 xlColorIndexNone 。*/
function test() {
    console.log(Range("B4").DisplayFormat.Borders.ColorIndex == xlColorIndexNone)
}
```

```javascript
/*此示例显示第一张工作表中 D6 单元格边框的粗细。*/
function test() {
    console.log(Worksheets.Item(1).Range("D6").DisplayFormat.Borders.Weight)
}
```
