---
title: "ListIndex 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/ListIndex%20%E5%B1%9E%E6%80%A7.html
---
# ListIndex (属性)

返回或设置列表框或组合框中当前选定项的索引号。**Long** 类型，可读写。

## 说明

不能将该属性应用于多选列表框。

## 示例

```javascript
/*本示例删除列表框中选定的数据项。如果 `Shapes.Item(2)` 代表的不是列表框，则本示例将无效。*/
function test() {
    let controlFormat = Worksheets.Item(1).Shapes.Item(2).ControlFormat
    controlFormat.RemoveItem(controlFormat.ListIndex)
}
```

```javascript
/*本示例显示活动工作表的第一个形状（列表框）当前选定项的索引号。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    console.log(shape.ControlFormat.ListIndex)
}
```
