---
title: "GradientVariant 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/GradientVariant%20%E5%B1%9E%E6%80%A7.html
---
# GradientVariant (属性)

以整数值（1 至 4）形式返回指定填充的底纹变量。此属性的值对应于\*\*“填充效果”**对话框中**“渐变”\*\*选项卡的渐变变量（从左到右、从上到下编号）。**Long** 型，只读。

## 说明

此属性为只读。使用 **OneColorGradient** 或 **TwoColorGradient**方法可设置填充的渐变变量。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状填充的底纹变量。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    console.log(fillFormat.GradientVariant)
}
```

```javascript
/*本示例将活动工作表中第二个形状填充的底纹变量赋值给B1单元格。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    Range("B1").Value2 = fillFormat.GradientVariant
}
```
