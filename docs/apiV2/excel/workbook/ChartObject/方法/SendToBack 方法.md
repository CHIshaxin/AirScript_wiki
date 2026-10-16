---
title: "SendToBack 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E6%96%B9%E6%B3%95/SendToBack%20%E6%96%B9%E6%B3%95.html
---
# SendToBack (方法)

将对象放到 z-次序的后面。

## 示例

```javascript
/*此示例将工作表 Sheet1 上第一个图表放到 z-次序的后面。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.SendToBack()
}
```

```javascript
/*此示例判断如果工作表 Sheet1 上图表 Chart 2 的索引为2，则将该图表放到 z-次序的后面。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    if (chartobject.Index == 2) {
        chartobject.SendToBack()
    }
}
```
