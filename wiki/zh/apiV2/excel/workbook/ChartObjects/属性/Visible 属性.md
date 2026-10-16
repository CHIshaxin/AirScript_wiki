---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: j2vduXaR
---

# Visible (属性)

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*此示例显示工作表 Sheet1 上所有图表是否可见。*/
function test() {
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    console.log(chartobjects.Visible)
}
```

```javascript
/*此示例设置图表工作表 Chart1 上所有图表为不可见。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    chartobjects.Visible = false
}
```
