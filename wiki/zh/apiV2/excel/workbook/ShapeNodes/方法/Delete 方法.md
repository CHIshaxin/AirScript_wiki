---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeNodes/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: E3WkOE4c
---

# Delete (方法)

删除对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int | 必填 | 要删除对象的索引。 |

## 示例

```javascript
/*本示例删除活动工作表中第一个形状的第二个节点。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.Item(1).Nodes.Delete(2)
}
```

```javascript
/*本示例删除第一张工作表中第二个形状的最后一个节点。*/
function test() {
    let nodes = Application.Worksheets.Item(1).Shapes.Item(2).Nodes
    nodes.Delete(nodes.Count)
}
```
