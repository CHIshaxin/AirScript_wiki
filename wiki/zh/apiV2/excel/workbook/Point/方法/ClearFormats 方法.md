---
title: "ClearFormats (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CQSHmo5z
---

# ClearFormats (方法)

清除对象的格式设置。

## 返回值

Variant

## 示例

```javascript
/*此示例清除图表工作表 Chart1 上第一个系列的第一个数据点的格式*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(1)
    point.ClearFormats()
}
```

```javascript
/*此示例清除图表工作表 Chart1 上第二个系列的所有奇数位的数据点的格式*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    for (let i = 1; i <= series.Points().Count; i += 2) {
        series.Points(i).ClearFormats()
    }
}
```
