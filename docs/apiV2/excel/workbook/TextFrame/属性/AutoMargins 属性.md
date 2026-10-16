---
title: "AutoMargins 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/AutoMargins%20%E5%B1%9E%E6%80%A7.html
---
# AutoMargins (属性)

返回或设置 ET 是否自动计算文本框边距。可读/写。

## 说明

如果 ET 自动计算文本框边距，则为 **True**；否则为 **False**。当此属性为 **True** 时，**MarginLeft**, **MarginRight**、**MarginTop**和 **MarginBottom**属性将被忽略。

## 返回值

Boolean

## 示例

```javascript
/*本示例将活动工作表中第二个标签窗体控件设置成能自动计算文本框边距。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame
    textFrame.AutoMargins = true
}
```

```javascript
/*本示例显示第一张工作表中第一个标签窗体控件是否自动计算文本框边距。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.Item(1).TextFrame
    console.log(textFrame.AutoMargins)
}
```
