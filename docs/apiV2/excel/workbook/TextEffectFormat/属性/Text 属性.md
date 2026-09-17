---
title: "Text 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
---
# Text (属性)

返回或设置指定对象中的文本。**String** 型，可读写。

## 示例

```javascript
/*本示例将第一张工作表中第三个艺术字的文本设置成“示例文本”。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(3)
    shape.TextEffect.Text = "示例文本"
}
```

```javascript
/*此示例显示活动工作表的第四个艺术字的文本。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(4)
    console.log(shape.TextEffect.Text)
}
```
