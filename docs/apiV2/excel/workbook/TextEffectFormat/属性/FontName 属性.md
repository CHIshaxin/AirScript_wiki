---
title: "FontName 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/FontName%20%E5%B1%9E%E6%80%A7.html
---
# FontName (属性)

返回或设置指定艺术字的字体名称。**String** 类型，可读写。

## 示例

```javascript
/*本示例将第一张工作表上第三个形状的字体名称设置为 Courier New（如果形状为艺术字）。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(3)
    if (shape.Type == msoTextEffect) {
        shape.TextEffect.FontName = "Courier New"
    }
}
```

```javascript
/*此示例显示活动工作表的第二个艺术字的字体名称。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.TextEffect.FontName)
}
```
