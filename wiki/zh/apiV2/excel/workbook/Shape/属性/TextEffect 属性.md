---
title: "TextEffect (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/TextEffect%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: rhqo9kH6
---

# TextEffect (属性)

该属性返回一个 **[TextEffectFormat](./../../TextEffectFormat/TextEffectFormat 对象.html)** 对象，它包含指定形状的文本效果格式属性。

## 示例

```javascript
/*此示例中，将第一张工作表中第一个形状的文本效果的字体设置为斜体。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(1)
    shape.TextEffect.FontItalic = msoTrue
}
```

```javascript
/*此示例中，将活动工作表中第三个形状的文本效果的字号设置为16。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    shape.TextEffect.FontSize = 16
}
```
