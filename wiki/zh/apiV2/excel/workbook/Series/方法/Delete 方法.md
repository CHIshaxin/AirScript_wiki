---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B-5SUA5n
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*此示例删除活动工作表上的第一个图表的第一个系列。*/
function test() {
    ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).Delete()
}
```

```javascript
/*此示例将图表工作表 Chart1 上的第二个数据系列删除。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    series2.Delete()
}
```
