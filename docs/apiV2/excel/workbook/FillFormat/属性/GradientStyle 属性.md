---
title: "GradientStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/GradientStyle%20%E5%B1%9E%E6%80%A7.html
---
# GradientStyle (属性)

返回指定填充的渐变样式。**MsoGradientStyle**，只读。

## 示例

```javascript
/*本示例显示活动工作表中第一个形状填充的渐变样式是否为混和。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(1).Fill
    console.log(fillFormat.GradientStyle == msoGradientMixed)
}
```

```javascript
/*本示例判断第一张工作表中第二个形状填充的渐变样式。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    if (fillFormat.GradientStyle == msoGradientFromCenter) {
        console.log("从中心到各个角的渐变")
    } else {
        console.log("msoGradientFromCenter之外的其它渐变")
    }
}
```
