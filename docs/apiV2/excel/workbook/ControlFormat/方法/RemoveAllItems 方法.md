---
title: "RemoveAllItems 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E6%96%B9%E6%B3%95/RemoveAllItems%20%E6%96%B9%E6%B3%95.html
---
# RemoveAllItems (方法)

删除 ET 列表框或组合框中的所有数据项。

## 示例

```javascript
/*本示例删除一个列表框中的所有数据项。如果 `Shapes.Item(2)` 代表的不是列表框，则本示例将无效。*/
function test() {
    Worksheets.Item(1).Shapes.Item(2).ControlFormat.RemoveAllItems()
}
```

```javascript
/*本示例删除活动工作表上第一个形状（组合框）的所有数据项。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.ControlFormat.RemoveAllItems()
}
```
