---
title: "SummaryColumn (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Outline/%E5%B1%9E%E6%80%A7/SummaryColumn%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BA45O4WJ
---

# SummaryColumn (属性)

返回或设置分级显示中的汇总列的位置。 **[XlSummaryColumn](./../../枚举/XlSummaryColumn 枚举.html)** 类型，可读写。

## 说明

| XlSummaryColumn 可为以下 常量之一。 |
| --- |
| xlSummaryOnRight 分级显示中的汇总列位于明细数据列的右侧。 |
| xlSummaryOnLeft 分级显示中的汇总列位于明细数据列的左侧。 |

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
/*本示例显示活动工作表中分级显示中的汇总列的位置是否位于明细数据列的左侧。*/
function test() {
    console.log(ActiveSheet.Outline.SummaryColumn == xlSummaryOnLeft)
}
```
