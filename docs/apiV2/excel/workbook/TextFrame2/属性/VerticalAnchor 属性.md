---
title: "VerticalAnchor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/VerticalAnchor%20%E5%B1%9E%E6%80%A7.html
---
# VerticalAnchor (属性)

返回或设置指定文本的垂直定位类型。可读/写 **MsoVerticalAnchor** 类型。

## 示例

```javascript
/*本示例显示活动工作表中第二个形状的文本框内文本是否为垂直居中。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame2
    console.log(textFrame.VerticalAnchor == msoAnchorMiddle)
}
```

```javascript
/*本示例使第一张工作表中第一个形状的文本与文本框的底部对齐。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2.VerticalAnchor = msoAnchorBottom
}
```
