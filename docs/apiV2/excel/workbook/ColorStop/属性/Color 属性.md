---
title: "Color 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStop/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
---
# Color (属性)

返回或设置指定对象的颜色。可读/写。

## 返回值

Variant

## 示例

```javascript
/*本示例显示活动单元格的第二个ColorStop对象颜色是否为蓝色。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Item(2)
    console.log(colorStop.Color == RGB(0, 0, 255))
}
```

```javascript
/*本示例将当前选定区域的第一个ColorStop对象的颜色设置为红色。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Item(1)
    colorStop.Color = RGB(255, 0, 0)
}
```
