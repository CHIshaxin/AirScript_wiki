---
title: "Series (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/Series%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: W5s_6_7F
---

# Series (对象)

代表图表上的系列。

## 说明

**[Series](./../Series/Series 对象.html)** 对象是 **[SeriesCollection](./../SeriesCollection/SeriesCollection 对象.html)** 集合的成员。

使用 **SeriesCollection**(*index*)（其中 *index*是系列索引号或名称）可返回一个 **[Series](./../Series/Series 对象.html)** 对象。下例设置 Sheet1 上嵌入式图表一中第一个系列的内部颜色。

系列索引号指明了系列添加到图表中的顺序。`SeriesCollection(1)` 是第一个添加到图表中的系列，而 `SeriesCollection(SeriesCollection.Count)`是最后一个添加到图表中的系列。

## 示例

```javascript
/*下面的示例设置工作表 Sheet1 上嵌入的第一个图表中第一个数据系列的内部颜色。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第二个数据系列的名称。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.Name)
}
```
