---
title: "TextEffect 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/TextEffect%20%E5%B1%9E%E6%80%A7.html
---
# TextEffect (属性)

该属性返回一个 **[TextEffectFormat](./../../TextEffectFormat/TextEffectFormat 对象.html)** 对象，它包含指定形状的文本效果格式属性。

## 示例

```javascript
/*此示例中，在第一张工作表中创建shpRange对象，将对象的文本效果的字体设置为斜体。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1,3])
    shpRange.TextEffect.FontItalic = msoTrue
}
```

```javascript
/*此示例中，在活动工作表中创建shpRange对象，将对象的文本效果的字号设置为16。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1])
    shpRange.TextEffect.FontSize = 16
}
```
