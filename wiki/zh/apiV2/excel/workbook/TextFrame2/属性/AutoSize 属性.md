---
title: "AutoSize (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/AutoSize%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: iKYr1IoI
---

# AutoSize (属性)

指定的对象的大小，可自动调整大小以适应其中包含的文字。可读/写 **MsoAutoSize**。

## 示例

```javascript
/*本示例使第一张工作表中第一个形状根据文本调整形状。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2.AutoSize = msoAutoSizeShapeToFitText
}
```

```javascript
/*本示例显示活动工作表中第二个形状是否根据形状调整文本。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame2
    console.log(textFrame.AutoSize == msoAutoSizeTextToFitShape)
}
```
