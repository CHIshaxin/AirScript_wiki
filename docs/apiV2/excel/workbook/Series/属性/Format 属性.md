---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例将图表工作表 Chart1 上的第二个数据系列的颜色设置为红色。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    series2.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*此示例将活动工作表上的第一个图表的第三个系列阴影向右侧偏移 5 磅。*/
function test() {
    ActiveSheet.ChartObjects(1).Chart.SeriesCollection(3).Format.Shadow.OffsetX = 5
}
```
