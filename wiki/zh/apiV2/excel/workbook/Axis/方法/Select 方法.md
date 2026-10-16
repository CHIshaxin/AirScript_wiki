---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: IB2ZDVRT
---

# Select (方法)

选择对象。

## 示例

```javascript
/*本示例选中工作表 Sheet1 中第一张图表的数值轴。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.Select()
}
```

```javascript
/*本示例选中图表工作表 Chart1 中图表的类别轴。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.Select()
}
```
