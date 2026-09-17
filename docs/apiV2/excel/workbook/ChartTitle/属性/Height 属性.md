---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回对象的高度（以磅为单位）。只读。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上图表 Chart 2 的标题的高度是否为 19.6 磅。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects("Chart 2").Chart.ChartTitle
    console.log(charttitle.Height == 19.6)
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的标题的高度。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    console.log(`图表工作表 Chart1 的标题高度：${charttitle.Height} 磅`)
}
```
