---
title: "SoftEdge (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/SoftEdge%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C-ZOYGvx
---

# SoftEdge (属性)

返回指定形状的 **SoftEdgeFormat** 对象，该对象包含形状的柔化边缘格式设置属性。只读。

## 示例

```javascript
/*本示例将活动工作表的第一个形状的柔化边缘效果设置为5磅。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.SoftEdge.Radius = 5
}
```

```javascript
/*本示例遍历第一张工作表的所有形状并设置它们的柔化边缘效果为10磅。*/
function test() {
    let shapes = ActiveSheet.Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        shapes.Item(i).SoftEdge.Radius = 10
    }
}
```
