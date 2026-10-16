---
title: "GradientColorType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/GradientColorType%20%E5%B1%9E%E6%80%A7.html
---
# GradientColorType (属性)

返回指定填充的渐变颜色类型。**MsoGradientColorType**，只读。

## 说明

**MsoGradientColorMixed** 是一个返回值，它仅指明指定区域中其他状态的组合。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient**方法可设置填充的渐变类型。

## 示例

```javascript
/*此示例判断第一张工作表第一个形状填充的渐变类型。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    if (fillFormat.GradientColorType == msoGradientColorMixed) {
        console.log("混合渐变")
    } else {
        console.log("其它渐变")
    }
}
```

```javascript
/*本示例显示活动工作表中第二个形状填充的渐变类型是否为单色渐变。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    console.log(fillFormat.GradientColorType == msoGradientOneColor)
}
```
