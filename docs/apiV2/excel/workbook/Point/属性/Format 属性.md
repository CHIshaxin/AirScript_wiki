---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例设置Chart1中第一个系列第三个数据点的填充颜色和边框线大小。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(3)
    point.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    point.Format.Line.Weight = 5
}
```

```javascript
/*此示例设置图表工作表 Chart1 上第二个系列的所有奇数位的数据点边框颜色*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    for (let i = 1; i <= series.Points().Count; i += 2) {
        series.Points(i).Format.Line.ForeColor.RGB = RGB(0, 0, 255)
    }
}
```
