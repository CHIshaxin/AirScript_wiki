---
title: "ObjectThemeColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorFormat/%E5%B1%9E%E6%80%A7/ObjectThemeColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BhDnpFQQ
---

# ObjectThemeColor (属性)

返回或设置一种映射到主题配色方案的颜色。可读/写 **MsoThemeColorIndex**。

## 示例

```javascript
/*本示例显示第一张工作表上第二个形状背景填充色映射到主题配色方案的颜色是否为无主题颜色。*/
function test() {
    let colorFormat = Worksheets.Item(1).Shapes.Item(2).Fill.BackColor
    console.log(colorFormat.ObjectThemeColor == msoNotThemeColor)
}
```

```javascript
/*本示例设置活动工作表中第一个形状前景填充色映射到主题配色方案的颜色。*/
function test() {
    ActiveSheet.Shapes.Item(1).Fill.ForeColor.ObjectThemeColor = msoThemeColorDark2
}
```
