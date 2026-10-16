---
title: "AddTextbox 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/AddTextbox%20%E6%96%B9%E6%B3%95.html
---
# AddTextbox (方法)

创建一个文本框。返回一个代表新文本框的 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Orientation | Kso.MsoTextOrientation | 必填 | 文本框的方向。 |
| Left | float | 必填 | 文本框左上角相对于文档左上角的位置（以磅为单位）。 |
| Top | float | 必填 | 文本框左上角相对于文档顶部的位置（以磅为单位）。 |
| Width | float | 必填 | 文本框的宽度（以磅为单位）。 |
| Height | float | 必填 | 文本框的高度（以磅为单位）。 |

## 返回值

Shape

## 示例

```javascript
/*本示例将包含文本“Test Box”的文本框添加到第一张工作表中。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let textbox = shapes.AddTextbox(msoTextOrientationHorizontal, 100, 10, 200, 50)
    textbox.TextFrame.Characters().Text = "Test Box"
}
```

```javascript
/*本示例将带有红色边框的文本框添加到活动工作表中。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let textbox = shapes.AddTextbox(msoTextOrientationVertical, 100, 10, 200, 50)
    textbox.Line.Weight = 2
    textbox.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
