---
title: "GradientDegree 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/GradientDegree%20%E5%B1%9E%E6%80%A7.html
---
# GradientDegree (属性)

以浮点数值的方式返回单色阴影填充的渐变程度，数值大小介于 0.0（暗）到 1.0（亮）之间。**Single**型，只读。

## 说明

此属性为只读。使用 **OneColorGradient**方法可设置填充的渐变程度。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状单色阴影填充的渐变程度。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    console.log(fillFormat.GradientDegree)
}
```

```javascript
/*本示例将活动工作表中第二个形状单色阴影填充的渐变程度赋值给A1单元格。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    Range("A1").Value2 = fillFormat.GradientDegree
}
```
