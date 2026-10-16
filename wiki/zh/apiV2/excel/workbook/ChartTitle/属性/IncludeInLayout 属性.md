---
title: "IncludeInLayout (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/IncludeInLayout%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BszPyKcG
---

# IncludeInLayout (属性)

如果在确定图表布局时图表标题将占用图表布局空间，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

## 说明

此属性对于图表是否处于自动版式模式无影响。如果用户使用**“图表上方”**命令添加标题，则图表将变得较小。如果用户随后删除标题，或者选择一个覆盖标题选项，则图表将变得较大，就好像标题不在图表上一样

## 示例

```javascript
/*本示例判断图表工作表 Chart1 的标题是否占用图表布局空间，并通知用户。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    if (charttitle.IncludeInLayout) {
        console.log("图表标题占用图表布局空间")
    } else {
        console.log("图表标题未占用图表布局空间")
    }
}
```

```javascript
/*本示例将工作表 Sheet1 上第一个图表的标题设置为不占用图表布局空间。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    charttitle.IncludeInLayout = false
}
```
