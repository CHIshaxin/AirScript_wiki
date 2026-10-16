---
title: "HorizontalAnchor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/HorizontalAnchor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: nps307uW
---

# HorizontalAnchor (属性)

返回或设置指定文本的水平定位类型。可读/写 **MsoHorizontalAnchor** 类型。

## 示例

```javascript
/*本示例将活动工作表中第一个文本框的文本设置成文本水平居中。*/
function test() {
    let textFrame2 = Application.ActiveSheet.Shapes.Item(1).TextFrame2
    textFrame2.HorizontalAnchor = msoAnchorCenter
}
```

```javascript
/*本示例显示第一张工作表中第二个形状的文本框的文本水平对齐方式是否是不对齐。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.Item(2).TextFrame2
    console.log(textFrame2.HorizontalAnchor == msoAnchorNone)
}
```
