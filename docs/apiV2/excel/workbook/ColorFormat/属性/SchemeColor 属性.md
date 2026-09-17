---
title: "SchemeColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorFormat/%E5%B1%9E%E6%80%A7/SchemeColor%20%E5%B1%9E%E6%80%A7.html
---
# SchemeColor (属性)

返回或设置一个 **Integer** 值，它代表 **Color** 对象的颜色，作为当前配色方案中的索引。

## 示例

```javascript
/*本示例设置活动工作表中第一个形状前景填充色的当前配色方案。*/
function test() {
    ActiveSheet.Shapes.Item(1).Fill.ForeColor.SchemeColor = 2
}
```

```javascript
/*本示例显示第一张工作表上第二个形状背景填充色的当前配色方案是否为绿色。*/
function test() {
    let colorFormat = Worksheets.Item(1).Shapes.Item(2).Fill.BackColor
    console.log(colorFormat.SchemeColor == 3)
}
```
