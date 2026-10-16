---
title: "ForeColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/ForeColor%20%E5%B1%9E%E6%80%A7.html
---
# ForeColor (属性)

返回或设置一个 **[ColorFormat](./../../ColorFormat/ColorFormat 对象.html)** 对象，它代表指定的前景填充色或纯色。

## 示例

```javascript
/*本示例将活动工作表中第三个形状阴影的前景色设置为红色。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(3).Shadow
    shadow.ForeColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*本示例将第一张工作表中第二个形状阴影的前景色设置为蓝色。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(2).Shadow
    shadow.ForeColor.RGB = RGB(0, 0, 255)
}
```
