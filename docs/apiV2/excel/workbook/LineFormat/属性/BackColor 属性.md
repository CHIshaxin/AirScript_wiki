---
title: "BackColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/BackColor%20%E5%B1%9E%E6%80%A7.html
---
# BackColor (属性)

返回或设置一个 **[ColorFormat](./../../ColorFormat/ColorFormat 对象.html)** 对象，它代表指定的填充背景色。

## 示例

```javascript
/*本示例将活动工作表的第一个形状线条的背景色设置为蓝色。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    line.BackColor.RGB = RGB(0, 0, 255)
}
```

```javascript
/*本示例显示第一张工作表中第二个形状线条的背景色是否为绿色。*/
function test() {
    let line = Worksheets.Item(1).Shapes.Item(2).Line
    console.log(line.BackColor.RGB == RGB(0, 255, 0))
}
```
