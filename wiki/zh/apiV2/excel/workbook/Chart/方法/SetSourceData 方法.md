---
title: "SetSourceData (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/SetSourceData%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B1V-MElU
---

# SetSourceData (方法)

为指定图表设置源数据区域。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Source | Range | 必填 | 包含源数据的区域。 |
| PlotBy | any | 可选 | 指定数据绘制方式。可为以下 XlRowCol 常量之一：xlColumns 或 xlRows。 |

## 示例

```javascript
/*本示例为图表工作表 Chart1 设置用户给定的工作表Sheet1区域A1:A10中的数据为源数据区域。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SetSourceData(Sheets.Item("Sheet1").Range("A1:A10"), xlColumns)
}
```

```javascript
/*本示例激活 Sheet1 工作表，并为B1等单元格分别赋值，然后设置区域B1:E1为第一张内嵌图表的源数据区域。*/
function test() {
    Application.Sheets.Item("Sheet1").Activate()
    Range("B1").Value2 = 5
    Range("C1").Value2 = 1
    Range("D1").Value2 = 7
    Range("E1").Value2 = 10
    let chart = Application.ActiveSheet.ChartObjects(1).Chart
    chart.SetSourceData(Range("B1:E1"), xlRows)
}
```
