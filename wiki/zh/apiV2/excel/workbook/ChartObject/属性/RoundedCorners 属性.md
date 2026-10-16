---
title: "RoundedCorners (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/RoundedCorners%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C60dWRI5
---

# RoundedCorners (属性)

如果嵌入式图表使用圆角，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例为工作表 Sheet1 中第一个图表添加圆角。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.RoundedCorners = true
}
```

```javascript
/*此示例显示第一张工作表中图表 Chart 2 是否使用圆角。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects("Chart 2")
    console.log(chartobject.RoundedCorners)
}
```
