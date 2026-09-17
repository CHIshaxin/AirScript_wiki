---
title: "Selection 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Selection%20%E5%B1%9E%E6%80%A7.html
---
# Selection (属性)

为 **[Application](./../../Application/Application 对象.html)** 对象返回在活动窗口中选定的对象。

## 说明

返回的对象类型取决于当前所选内容（例如，如果选择了单元格，此属性将返回 Range 对象）。如果未选择任何内容，**Selection** 属性将返回 **null**。

在不使用对象识别符的情况下，使用此属性等效于使用 `Application.Selection`。

## 示例

```javascript
/*本示例清空 Sheet1 的选定对象（假定选定对象为单元格区域）。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Selection.Clear()
}
```

```javascript
/*本示例显示选定对象的 wpsjs 加载项 对象类型。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    console.log("The selection object type is " + Application.Selection.constructor.name)
}
```
