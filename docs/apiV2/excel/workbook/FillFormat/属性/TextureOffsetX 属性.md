---
title: "TextureOffsetX 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/TextureOffsetX%20%E5%B1%9E%E6%80%A7.html
---
# TextureOffsetX (属性)

返回指定填充的偏移 X 值。可读/写 **Single** 类型。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状填充的偏移 X 值。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    console.log(fillFormat.TextureOffsetX)
}
```

```javascript
/*本示例将活动工作表中第二个形状填充的偏移 X 值设置为15。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    fillFormat.TextureOffsetX = 15
}
```
