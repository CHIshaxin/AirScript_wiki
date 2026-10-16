---
title: "Workbook 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/Workbook%20%E5%AF%B9%E8%B1%A1.html
---
# Workbook (对象)

代表一个 ET 工作簿。

## 说明

**[Workbook](./../Workbook/Workbook 对象.html)** 对象是 Workbooks 集合的成员。**Workbooks** 集合包含 ET 中当前打开的所有 **[Workbook](./../Workbook/Workbook 对象.html)** 对象。

## 示例

使用 **Workbooks**(*index*)（其中 *index* 是工作簿名称或索引号）可返回一个 **[Workbook](./../Workbook/Workbook 对象.html)** 对象。下例激活工作簿一。

```javascript
/*以下示例激活第一个工作簿。*/
function test() {
    Application.Workbooks.Item(1).Activate()
}
```

编号指示创建或打开工作簿的顺序。`Workbooks(1)` 是创建的第一个工作簿，而 `Workbooks(Workbooks.Count)`Workbooks 是最后一个。激活某工作簿并不更改其索引号。所有工作簿均包括在索引计数中，即便是隐藏工作簿也是如此。

Name 属性返回工作簿名称。您不能通过使用此属性来设置该名称；如果您需要更改该名称，请使用 SaveAs 方法，将该工作簿保存为其他名称。

下例激活名为“Cogs.xls”的工作簿（该工作簿必须已经在 ET 中打开）中的 Sheet1。

Workbooks("Cogs.xls").Worksheets("Sheet1").Activate()

**ActiveWorkbook**属性返回当前处于活动状态的工作簿。下例设置活动工作簿作者的名称。

```javascript
/*以下示例设置活动工作簿作者的名称。*/
function test() {
    Application.ActiveWorkbook.Author = "Jean Selva"
}
```
