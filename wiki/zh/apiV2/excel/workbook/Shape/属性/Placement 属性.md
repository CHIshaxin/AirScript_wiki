---
title: "Placement (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Placement%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D4-8e_QB
---

# Placement (属性)

返回或设置一个 **[XlPlacement](./../../枚举/XlPlacement 枚举.html)** 值，它代表对象附加到对象下面的单元格的方式。

## 示例

```javascript
/*本示例设置活动工作表的第一个形状为可自由浮动（既不随下方单元格移动，也不随其改变大小）。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.Placement = xlFreeFloating
}
```

```javascript
/*本示例设置第一个工作表的第一个形状随单元格移动和调整大小。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(1)
    shape.Placement = xlMoveAndSize
}
```
