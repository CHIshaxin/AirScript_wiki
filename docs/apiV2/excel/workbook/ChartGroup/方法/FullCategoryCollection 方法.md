---
title: "FullCategoryCollection 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E6%96%B9%E6%B3%95/FullCategoryCollection%20%E6%96%B9%E6%B3%95.html
---
# FullCategoryCollection (方法)

返回一个对象，该对象代表图表组中 CategoryCollection 集合（所有可见类别和筛选类别的集合）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 类别的名称或编号。 |

## 示例

```javascript
/*此示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组的第一个类别是否筛选出序列。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.FullCategoryCollection(1).IsFiltered)
}
```

```javascript
/*此示例显示图表工作表 Chart1 第一个图表组的类别数量。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.FullCategoryCollection().Count)
}
```
