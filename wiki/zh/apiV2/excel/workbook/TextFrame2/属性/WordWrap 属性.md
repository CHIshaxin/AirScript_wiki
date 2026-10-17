---
title: "WordWrap (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/WordWrap%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CfP4V8iT
---

# WordWrap (属性)

返回或设置形状边界内外的文本换行。可读/写 **MsoTriState** 类型。

## 示例

```javascript
/*本示例显示活动工作表中第二个形状的文本是否支持自动换行。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame2
    console.log(textFrame.WordWrap == msoTrue)
}
```

```javascript
/*本示例将第一张工作表中第一个形状的文本设置为不支持自动换行。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2.WordWrap = msoFalse
}
```
