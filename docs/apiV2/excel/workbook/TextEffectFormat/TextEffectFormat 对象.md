---
title: "TextEffectFormat 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/TextEffectFormat%20%E5%AF%B9%E8%B1%A1.html
---
# TextEffectFormat (对象)

包含应用于艺术字对象的属性和方法。

## 说明

使用 **TextEffect** 属性可返回一个 **[TextEffectFormat](./../TextEffectFormat/TextEffectFormat 对象.html)** 对象。

## 示例

```javascript
/*下例为第一张工作表上的形状一设置字体名称及格式。要运行本示例，形状一必须是艺术字对象。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let TextEffect = worksheet.Shapes(1).TextEffect
    TextEffect.FontName = "Courier New"
    TextEffect.FontBold = true
    TextEffect.FontItalic = true
}
```

```javascript
/*此示例显示活动工作表的第三个艺术字的字号。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.TextEffect.FontSize)
}
```
