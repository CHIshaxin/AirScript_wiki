---
title: "TintAndShade 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStop/%E5%B1%9E%E6%80%A7/TintAndShade%20%E5%B1%9E%E6%80%A7.html
---
# TintAndShade (属性)

返回或设置指定对象的淡色和阴影。可读/写。

## 返回值

Variant

## 示例

```javascript
/*对当前选定内容应用淡色和阴影。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    colorStop.ThemeColor = xlThemeColorAccent1
    colorStop.TintAndShade = 0
}
```

```javascript
/*本示例显示活动单元格第二个ColorStop对象的TintAndShade属性是否为1。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Item(3)
    console.log(colorStop.TintAndShade == 1)
}
```
