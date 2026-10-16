---
title: "ActiveSheet 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/ActiveSheet%20%E5%B1%9E%E6%80%A7.html
---
# ActiveSheet (属性)

返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **null**。

## 说明

如果不指定对象识别符，则此属性返回活动工作簿中的活动工作表。

如果某个工作簿出现在若干个窗口中，那么该工作簿的 **ActiveSheet** 属性在不同窗口中可能不同。

## 示例

```javascript
/*此示例显示活动工作表的名称。*/
function test() {
    console.log("The name of the active sheet is " + Application.ActiveSheet.Name)
}
```

```javascript
/*此示例将活动工作表上单元格 A1 的值设置为 3.14159。*/
function test() {
    Application.ActiveSheet.Range("A1").Value2 = 3.14159
}
```
