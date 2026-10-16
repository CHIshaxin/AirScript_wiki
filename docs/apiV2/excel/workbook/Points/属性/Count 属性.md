---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Points/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*此示例显示第一张工作表上第一个图表的第二个数据系列中数据点的数量。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Points().Count)
}
```
