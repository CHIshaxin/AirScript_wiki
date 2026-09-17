---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheets/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
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
/*本示例在第一张工作表前新建1张图表。*/
function test() {
    Worksheets.Add(Worksheets.Item(1), null, 1, xlChart)
}
```

```javascript
/*本示例在第一张工作表后新建2张工作表，并将其中一张工作表命名为成绩表。*/
function test() {
    Worksheets.Add(null, Worksheets.Item(1), 2).Name = "成绩表"
}
```
