---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeNodes/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Integer** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示活动工作表中第一个形状的节点数量。*/
function test() {
    let shapes = ActiveSheet.Shapes
    console.log(shapes.Item(1).Nodes.Count)
}
```
