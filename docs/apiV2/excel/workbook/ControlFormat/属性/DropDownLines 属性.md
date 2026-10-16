---
title: "DropDownLines 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/DropDownLines%20%E5%B1%9E%E6%80%A7.html
---
# DropDownLines (属性)

返回或设置组合框中下拉区域中显示的列表行的数目。**Long** 类型，可读写。

## 示例

```javascript
/*本示例创建一个组合框，并将其下拉部分设置为 10 行。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.AddFormControl(xlDropDown, 10, 10, 100, 10)
    shape.ControlFormat.DropDownLines = 10
}
```

```javascript
/*本示例显示活动工作表的第二个形状（组合框）中下拉区域中显示的列表行的数目。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ControlFormat.DropDownLines)
}
```
