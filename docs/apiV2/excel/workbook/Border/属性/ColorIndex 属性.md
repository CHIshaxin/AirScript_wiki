---
title: "ColorIndex 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Border/%E5%B1%9E%E6%80%A7/ColorIndex%20%E5%B1%9E%E6%80%A7.html
---
# ColorIndex (属性)

返回或设置一个 **Variant** 值，它代表边框的颜色。

## 说明

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 XlColorIndex 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

## 示例

```javascript
/*本示例为 Chart1 中数值坐标轴的主要网格线设置颜色。*/
function test() {
    let axis = Application.Charts.Item("Chart1").Axes(xlValue)
    if (axis.HasMajorGridlines) {
        //set color to blue
        axis.MajorGridlines.Border.ColorIndex = 5
    }
}
```

```javascript
/*本示例显示 Chart1 中数值坐标轴的边框是否为无颜色。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.Axes(xlValue).Border.ColorIndex == xlColorIndexNone)
}
```
