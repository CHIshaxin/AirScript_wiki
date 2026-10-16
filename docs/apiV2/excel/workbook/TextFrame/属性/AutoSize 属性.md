---
title: "AutoSize 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/AutoSize%20%E5%B1%9E%E6%80%A7.html
---
# AutoSize (属性)

如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例使第一个形状的文本框能自动调整大小，以适应其中所包含的文字。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.AutoSize = true
}
```

```javascript
/*本示例显示活动工作表中第二个形状的文本框是否能自动调整大小，以适应其中所包含的文字。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame
    console.log(textFrame.AutoSize)
}
```
