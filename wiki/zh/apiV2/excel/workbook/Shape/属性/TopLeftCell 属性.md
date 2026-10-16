---
title: "TopLeftCell (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/TopLeftCell%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: vW46sB2m
---

# TopLeftCell (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表位于指定对象左上角下方的单元格。只读。

## 示例

```javascript
/*本示例将活动工作表中第一个形状左上角所在的单元格的值设置为 "Hello"。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    let range = shape.TopLeftCell
    range.Value2 = "Hello"
}
```

```javascript
/*本示例将第一张工作表中第二个形状左上角所在的单元格的背景色设置为红色。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(2)
    let range  = shape.TopLeftCell
    range.Interior.Color = RGB(255, 0, 0)
}
```
