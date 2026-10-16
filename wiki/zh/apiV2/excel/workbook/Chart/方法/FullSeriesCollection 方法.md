---
title: "FullSeriesCollection (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/FullSeriesCollection%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CeotfQGH
---

# FullSeriesCollection (方法)

启用检索 由 Index 参数指定的筛选出序列。

## 说明

除非用户在“选择数据”对话框中启用了“在隐藏行和列中显示数据”选项，否则隐藏行或列中的序列对象不会显示在当前系列集合中。

还可以在引号中使用系列名称：

expression.FullSeriesCollection("系列名称")

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 可选 | 筛选出的 Series 对象的索引号。 |

## 返回值

Object

## 示例

```javascript
/*本示例将图表工作表 Chart1 “系列1”设置为红色。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.FullSeriesCollection("系列1").Format.Fill.ForeColor.RGB = rgbRed
}
```

```javascript
/*本示例删除工作表 Sheet1 第一个内嵌图表的第二个系列。*/
function test() {
    let chart = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart
    chart.FullSeriesCollection(2).Delete()
}
```
