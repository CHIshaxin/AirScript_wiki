---
title: "SummaryRow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Outline/%E5%B1%9E%E6%80%A7/SummaryRow%20%E5%B1%9E%E6%80%A7.html
---
# SummaryRow (属性)

返回或设置分级显示中汇总行的位置。 **[XlSummaryRow](./../../枚举/XlSummaryRow 枚举.html)** 类型，可读写。

## 说明

| XlSummaryRow 可为以下常量之一。 |
| --- |
| xlSummaryBelow 分级显示中的汇总行位于明细数据行的下方。 |
| xlSummaryAbove 分级显示中的汇总行位于明细数据行的上方。 |

若要进行 WPS 风格的分级显示，请将 **[XlSummaryRow](./../../枚举/XlSummaryRow 枚举.html)** 设置为 **xlAbove**，使分类标题位于明细信息的上方。若要进行会计风格的分级显示，请将 **[XlSummaryRow](./../../枚举/XlSummaryRow 枚举.html)** 设置为 **xlBelow**，使求和位于明细信息的下方。

## 示例

```javascript
/*本示例创建一个使用自动样式的分级显示，汇总行位于明细数据行的上方，汇总列位于明细数据列的右侧。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    Selection.AutoOutline()
    let outline = ActiveSheet.Outline
    outline.SummaryRow = xlSummaryAbove
    outline.SummaryColumn = xlSummaryOnRight
    outline.AutomaticStyles = true
}
```

```javascript
/*本示例显示活动工作表中分级显示中的汇总行的位置是否位于明细数据行的下方。*/
function test() {
    console.log(ActiveSheet.Outline.SummaryRow == xlSummaryBelow)
}
```
