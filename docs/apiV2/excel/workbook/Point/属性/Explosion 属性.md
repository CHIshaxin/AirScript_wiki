---
title: "Explosion 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Explosion%20%E5%B1%9E%E6%80%A7.html
---
# Explosion (属性)

返回或设置饼图或圆环图的扇区的分离程度值。如果没有分离，则返回 0（零），即扇区的中心与饼图中心重合。**Long** 型，可读写。

## 示例

```javascript
/*此示例设置 Chart1 中第二个点的分离值。此示例应在饼图上运行。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2)
    point.Explosion = 20
}
```

```javascript
/*此示例判断活动工作表第一个图表上第二个数据点是否分离。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    let point = ActiveChart.SeriesCollection(1).Points(2)
    if (point.Explosion > 0) {
        console.log("数据点已分离")
    } else {
        console.log("数据点未分离")
    }
}
```
