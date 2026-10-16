---
title: "RemoveItem (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E6%96%B9%E6%B3%95/RemoveItem%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CnZ97I8F
---

# RemoveItem (方法)

从列表框或组合框中删除一个或多个数据项。

## 说明

如果指定对象定义了填充区域，则该方法无效。

使用 **RemoveAllItems** 方法删除 ET 列表框或组合框中的所有数据项。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int4 | 必填 | 第一个要删除的数据项编号。有效值为从 1 到列表中的数据项数目（**ListCount** 属性返回的值）。 |
| Count | any | 可选 | 要删除的数据项数目，从 *Index* 项开始。如果省略该参数，则只删除一个数据项。如果 *Index* + *Count* 超过列表中的数据项数目，则删除列表中从 *Index* 项到最后一项的所有数据项并且不报错。 |

## 示例

```javascript
/*本示例删除列表框中选定的数据项。如果 `Shapes.Item(2)` 代表的不是列表框，则本示例将无效。*/
function test() {
    let controlFormat = Worksheets.Item(1).Shapes.Item(2).ControlFormat
    controlFormat.RemoveItem(controlFormat.ListIndex)
}
```

```javascript
/*本示例删除活动工作表的第一个形状（列表框）的两个数据项，从索引1开始。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.ControlFormat.RemoveItem(1, 2)
}
```
