---
title: "ActiveCell 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/ActiveCell%20%E5%B1%9E%E6%80%A7.html
---
# ActiveCell (属性)

返回一个 Range 对象，该对象代表活动窗口（顶部窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。

## 说明

如果不指定对象识别符，此属性返回活动窗口中的活动单元格。

请仔细区分活动单元格和选定区域。活动单元格为选定区域内部的一个单元格。而选定区域可以包含多个单元格，但只有一个单元格为活动单元格。

下列表达式都是返回活动单元格，并且都是等效的。

```javascript
Application.ActiveCell
Application.ActiveWindow.ActiveCell
```

## 示例

```javascript
/*此示例在消息框中显示活动单元格的值。由于如果活动表不是工作表则 ActiveCell 属性无效，所以此示例使用 ActiveCell 属性之前先激活 Sheet1。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    console.log(Application.ActiveCell.Value2)
}
```

```javascript
/*此示例更改活动单元格的字体格式设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let font = Application.ActiveCell.Font
    font.Bold = true
    font.Italic = true
}
```
