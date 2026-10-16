---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: iVj6-94W
---

# Visible (属性)

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例将活动工作表中第二个形状的填充设置为不可见。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    fillFormat.Visible = msoFalse
}
```

```javascript
/*本示例将第一张工作表中第一个形状的填充设置为可见。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    fillFormat.Visible = msoTrue
}
```
