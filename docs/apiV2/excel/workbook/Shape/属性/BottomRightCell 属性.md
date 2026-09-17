---
title: "BottomRightCell 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/BottomRightCell%20%E5%B1%9E%E6%80%A7.html
---
# BottomRightCell (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表位于该对象右下角下方的单元格。只读。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状右下角下的单元格的值。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    console.log(shapes.Item(1).BottomRightCell.Value2)
}
```

```javascript
/*本示例显示活动工作表中第一个形状右下角下的单元格的地址，将形状高度缩小至一半，并再次显示形状右下角下的单元格的地址。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let range = shapes.Item(1).BottomRightCell
    console.log(range.Address())
    shapes.Item(1).Height = shapes.Item(1).Height / 2
    let newRange = shapes.Item(1).BottomRightCell
    console.log(newRange.Address())
}
```
