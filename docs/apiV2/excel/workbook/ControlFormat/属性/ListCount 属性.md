---
title: "ListCount 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/ListCount%20%E5%B1%9E%E6%80%A7.html
---
# ListCount (属性)

返回列表框或者组合框中数据项的数目。如果列表中没有数据项，则返回 0（零）。**Long** 类型，只读。

## 示例

```javascript
/*本示例对组合框进行调整，使其可以显示列表中的所有数据项。如果第一个形状不代表组合框，则本示例无效。*/
function test() {
    let controlFormat = Worksheets.Item(1).Shapes.Item(1).ControlFormat
    controlFormat.DropDownLines = controlFormat.ListCount
}
```

```javascript
/*本示例显示活动工作表的第二个形状（列表框）数据项的数目。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ControlFormat.ListCount)
}
```
