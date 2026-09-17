---
title: "LinkedCell 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/LinkedCell%20%E5%B1%9E%E6%80%A7.html
---
# LinkedCell (属性)

返回或设置指向控制值的工作表区域。如果为这些单元格赋值，则指定控制也会取得相应的值。与此类似，如果更改控制的值，则单元格的值也作相应变动。**String** 型，可读写。

## 说明

不能将该属性应用于多选列表框。

## 示例

```javascript
/*此示例为第一张工作表添加复选框，并将 A1 单元格设置为此复选框的附属单元格。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.AddFormControl(xlCheckBox, 10, 10, 100, 10)
    shape.ControlFormat.LinkedCell = "A1"
}
```

```javascript
/*本示例显示活动工作表的第三个形状（滚动条）链接单元格。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ControlFormat.LinkedCell)
}
```
