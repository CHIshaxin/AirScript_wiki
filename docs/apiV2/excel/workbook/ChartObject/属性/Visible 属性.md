---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*此示例判断工作表 Sheet1 上图表 Chart 2 是否可见，并通知用户。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    if (chartobject.Visible) {
        console.log("图表 Chart 2 可见")
    } else {
        console.log("图表 Chart 2 不可见")
    }
}
```

```javascript
/*此示例将工作表 Sheet1 上第一个图表设置为不可见。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Visible = false
}
```
