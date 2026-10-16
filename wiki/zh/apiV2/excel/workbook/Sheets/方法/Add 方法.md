---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 7I6hck5G
---

# Add (方法)

新建工作表、图表或宏表。新建的工作表将成为活动工作表。

## 说明

如果同时省略 *Before* 和 *After*，则新工作表插入到活动工作表之前。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Before | any | 可选 | 指定工作表的对象，新建的工作表将置于此工作表之前。 |
| After | any | 可选 | 指定工作表的对象，新建的工作表将置于此工作表之后。 |
| Count | any | 可选 | 要添加的工作表数。默认值为 1。 |
| Type | any | 可选 | 指定工作表类型。可以为下列 XlSheetType 常量之一：xlWorksheet、xlChart、xlExcel4MacroSheet 或 xlExcel4IntlMacroSheet。如果基于现有模板插入工作表，则指定该模板的路径。默认值为 xlWorksheet。 |

## 返回值

一个 Object 值，它代表新的工作表、图表或宏表。

## 示例

```javascript
/*本示例将新建工作表插入到活动工作簿的最后一张工作表之前。*/
function test() {
    Application.ActiveWorkbook.Sheets.Add(Application.Worksheets.Item(Worksheets.Count))
}
```

```javascript
/*本示例在活动工作簿中的最后一个工作表之后插入一个新图表，并将返回的对象赋值sheet变量。*/
function test() {
    let sheet = ActiveWorkbook.Sheets.Add(null, ActiveWorkbook.Worksheets.Item(Worksheets.Count), null, xlChart)
}
```
