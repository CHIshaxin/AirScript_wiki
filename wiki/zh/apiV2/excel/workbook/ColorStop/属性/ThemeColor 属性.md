---
title: "ThemeColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStop/%E5%B1%9E%E6%80%A7/ThemeColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: XgAgxlV8
---

# ThemeColor (属性)

返回或设置指定对象的主颜色。可读/写。

## 返回值

Long

## 示例

```javascript
/*对当前选定内容应用主题颜色。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    colorStop.ThemeColor = xlThemeColorAccent1
    colorStop.TintAndShade = 0
}
```

```javascript
/*本示例显示活动单元格第三个ColorStop对象的主题颜色是否为 xlThemeColorDark1。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Item(3)
    console.log(colorStop.ThemeColor == xlThemeColorDark1)
}
```
