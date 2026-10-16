---
title: "ClearFormats (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DHisj0x-
---

# ClearFormats (方法)

清除对象的格式设置。

## 示例

```javascript
/*此示例清除图表工作表 Chart1 上的第一个系列格式设置。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series1.ClearFormats()
}
```

```javascript
/*此示例删除活动工作表上的第一个图表的第二个系列的格式设置。*/
function test() {
    ActiveSheet.ChartObjects(1).Chart.SeriesCollection(2).ClearFormats()
}
```
