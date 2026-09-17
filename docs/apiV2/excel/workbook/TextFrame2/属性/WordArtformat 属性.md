---
title: "WordArtformat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/WordArtformat%20%E5%B1%9E%E6%80%A7.html
---
# WordArtformat (属性)

返回或设置指定文本框的艺术字类型。可读/写 **MsoPresetTextEffect** 类型。

## 示例

```javascript
/*本示例将第一张工作表中第一个形状的文本框艺术字类型设置成第十个文本效果。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2.WordArtformat = msoTextEffect10
}
```

```javascript
/*本示例显示活动工作表中第二个形状的文本框的艺术字类型是否为msoTextEffect1。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame2
    console.log(textFrame.WordArtformat == msoTextEffect1)
}
```
