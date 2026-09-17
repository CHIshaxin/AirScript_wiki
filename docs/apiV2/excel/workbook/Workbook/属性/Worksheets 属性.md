---
title: "Worksheets 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/Worksheets%20%E5%B1%9E%E6%80%A7.html
---
# Worksheets (属性)

返回一个 Worksheets 集合，它代表指定工作簿中所有工作表。 **[Worksheets](./../../Worksheets/Worksheets 对象.html)** 对象，只读。

## 说明

在不使用对象识别符的情况下，使用此属性将返回活动工作簿中所有的工作表。

此属性不返回宏表；使用 **Excel4MacroSheets** 属性或 **Excel4IntlMacroSheets** 属性可返回这些表。

## 示例

```javascript
/*此示例显示活动工作簿中 sheet1 上单元格 A1 中的值。*/
function test() {
    console.log(Application.Worksheets.Item("Sheet1").Range("A1").Value2)
}
```

```javascript
/*此示例显示活动工作簿中每个工作表的名称。*/
function test() {
    for (let i = 1; i <= Application.Worksheets.Count; i++) {
        console.log(Application.Worksheets.Item(i).Name)
    }
}
```

```javascript
/*此示例向活动工作簿添加新工作表，并设置该工作表的名称。*/
function test() {
    let newSheet = Application.Worksheets.Add()
    newSheet.Name = "current Budget"
}
```
