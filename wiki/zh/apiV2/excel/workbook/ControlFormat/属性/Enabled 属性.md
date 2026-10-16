---
title: "Enabled (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/Enabled%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C4S4CJ23
---

# Enabled (属性)

如果启用对象，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将第一张工作表的第一个形状设置成禁用。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ControlFormat.Enabled = false
}
```

```javascript
/*本示例显示活动工作表的第二个形状（组合框）是否启用对象。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ControlFormat.Enabled)
}
```
