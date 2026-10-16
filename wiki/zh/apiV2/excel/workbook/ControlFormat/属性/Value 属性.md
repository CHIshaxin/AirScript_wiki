---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dhcewc-A
---

# Value (属性)

返回或设置一个 **Long** 值，它代表指定的控件格式的值。

## 示例

```javascript
/*本示例将第一张工作表的第一个形状（组合框）的 Value 属性设置为1。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ControlFormat.Value = 1
}
```

```javascript
/*本示例显示活动工作表的第二个形状（列表框）选定的数据项的编号。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ControlFormat.Value)
}
```
