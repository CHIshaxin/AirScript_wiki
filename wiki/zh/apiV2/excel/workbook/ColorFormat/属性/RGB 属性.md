---
title: "RGB (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorFormat/%E5%B1%9E%E6%80%A7/RGB%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D6709YWv
---

# RGB (属性)

返回或设置一个 **Long** 值，它代表指定颜色的红绿蓝 (RGB) 值。

## 示例

```javascript
/*本示例显示第一张工作表上第二个形状背景填充色是否为红色。*/
function test() {
    let colorFormat = Worksheets.Item(1).Shapes.Item(2).Fill.BackColor
    console.log(colorFormat.RGB == RGB(255, 0, 0))
}
```

```javascript
/*本示例设置活动工作表中第一个形状前景填充色为绿色。*/
function test() {
    ActiveSheet.Shapes.Item(1).Fill.ForeColor.RGB = RGB(0, 255, 0)
}
```
