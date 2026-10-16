---
title: "Paste (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/Paste%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DUzJS1Cf
---

# Paste (方法)

将剪贴板中的图表数据粘贴到指定的图表中。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Type | any | 可选 | 指定要粘贴的图表信息（如果剪贴板中有一个图表）。可以为以下 XlPasteType 常量之一：xlPasteFormats、xlPasteFormulas 或 xlPasteAll。默认值为 xlPasteAll。如果剪贴板上的数据不是图表数据，则不能使用该参数。 |

## 示例

```javascript
/*本示例将工作表 Sheet1 上单元格区域 B1:B5 中的数据粘贴到 Chart1 中。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("B1:B5").Copy()
    Application.Charts.Item("Chart1").Paste()
}
```

```javascript
/*本示例假定用户已选中图表并复制，并将该图表粘贴到工作表Sheet1上的第一个内嵌图表中。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.Paste(xlPasteAll)
}
```
