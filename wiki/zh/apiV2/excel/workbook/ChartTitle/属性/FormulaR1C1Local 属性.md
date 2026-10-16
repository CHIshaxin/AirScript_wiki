---
title: "FormulaR1C1Local (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/FormulaR1C1Local%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BEPMbFVH
---

# FormulaR1C1Local (属性)

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*本示例设置工作表 Sheet1 上图表 Chart 1 的标题为该工作表中 C3 单元格的内容。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects("Chart 1").Chart.ChartTitle
    charttitle.FormulaR1C1Local = "='Sheet1'!R3C3"
}
```

```javascript
/*本示例设置工作表 Sheet1 中图表 Chart 2 的标题为该工作表中 A1 单元格的内容。*/
function test() {
    let charttitle = Application.Charts.Item("Sheet1").ChartObjects("Chart 2").Chart.ChartTitle
    charttitle.FormulaR1C1Local = "='Sheet1'!R1C1"
}
```
