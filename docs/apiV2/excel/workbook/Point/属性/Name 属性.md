---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回对象名称。只读。

## 说明

此属性对应于图表中数据点的名称。名称用下面的格式表示：S<*系列序号*>P<*数据标记序号*>。

## 返回值

String

## 示例

```javascript
/*此示例显示 Chart1 中第二个数据系列的第三个数据点的名称。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Points(3)
    console.log(`数据点的名称是 ${point.Name}`)
}
```

```javascript
/*此示例将 Chart1 中所有系列的数据点中名称为 "S3P3" 的背景色设置为红色。*/
function test() {
    let seriesColltion = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection()
    for (let i = 1; i <= seriesColltion.Count; i++) {
        for (let n = 1; n <= seriesColltion.Item(i).Points().Count; n++) {
            if (seriesColltion.Item(i).Points(n).Name == "S3P3") {
                seriesColltion.Item(i).Points(n).MarkerBackgroundColor = RGB(255, 0, 0)
            }
        }
    }
}
```
