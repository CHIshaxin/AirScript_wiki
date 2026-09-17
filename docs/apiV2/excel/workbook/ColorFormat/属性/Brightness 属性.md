---
title: "Brightness 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorFormat/%E5%B1%9E%E6%80%A7/Brightness%20%E5%B1%9E%E6%80%A7.html
---
# Brightness (属性)

返回或设置指定对象的发光度。可读/写。

## 说明

该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数字。

## 返回值

Float

## 示例

```javascript
/*本示例设置活动工作表中第一个形状的前景填充色的亮度。*/
function test() {
    Application.ActiveSheet.Shapes.Item(1).Fill.ForeColor.Brightness = 0.5
}
```

```javascript
/*本示例显示第一张工作表上第二个形状背景填充色的亮度。*/
function test() {
    let colorFormat = Worksheets.Item(1).Shapes.Item(2).Fill.BackColor
    console.log(colorFormat.Brightness)
}
```
