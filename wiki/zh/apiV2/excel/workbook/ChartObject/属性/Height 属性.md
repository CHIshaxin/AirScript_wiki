---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: YuooNPj1
---

# Height (属性)

返回或设置一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例将工作表 Sheet1 上嵌入的第一个图表的高度设置为 200 磅。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Height = 200
}
```

```javascript
/*本示例显示第一张工作表上嵌入的图表 Chart 2 的高度。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects("Chart 2")
    console.log(`图表 Chart 2 的高度为：${chartobject.Height} 磅`)
}
```
