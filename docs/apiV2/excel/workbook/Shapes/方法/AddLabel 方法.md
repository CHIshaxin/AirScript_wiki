---
title: "AddLabel 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/AddLabel%20%E6%96%B9%E6%B3%95.html
---
# AddLabel (方法)

创建一个连接符。返回一个代表新连接符的 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Orientation | Kso.MsoTextOrientation | 必填 | 标签中文本的方向。 |
| Left | float | 必填 | 标签左上角相对于文档左上角的位置（以磅为单位）。 |
| Top | float | 必填 | 标签左上角相对于文档顶部的位置（以磅为单位）。 |
| Width | float | 必填 | 标签的宽度（以磅为单位）。 |
| Height | float | 必填 | 标签的高度（以磅为单位）。 |

## 返回值

Shape

## 示例

```javascript
/*本示例将包含文本“Test Label”的垂直标签添加到第一张工作表中。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let label = shapes.AddLabel(msoTextOrientationVertical, 100, 100, 60, 150)
    label.TextFrame.Characters().Text = "Test Label"
}
```

```javascript
/*本示例将包含实线样式的水平标签添加到活动工作表中。*/
function test() {
    let label = ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 100, 100, 50, 50)
    label.Line.Weight = 2
    label.Line.DashStyle = msoLineSolid
}
```
