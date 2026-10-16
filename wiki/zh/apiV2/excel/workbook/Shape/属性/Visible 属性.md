---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CZzyCDXT
---

# Visible (属性)

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例将第一张工作表中的第二个形状设置为不可见。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(2)
    shape.Visible = msoFalse
}
```

```javascript
/*本示例将活动工作表上所有形状设置为可见。*/
function test() {
    let shapes = ActiveSheet.Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        shapes.Item(i).Visible = msoTrue
    }
}
```
