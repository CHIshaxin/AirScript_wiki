---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C_CwKlIZ
---

# Height (属性)

返回或设置一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*此示例将工作表 Sheet1 中所有图表的高度设置为工作表 1 到 15 行的距离。*/
function test() {
    let sheet = Application.Sheets.Item("Sheet1")
    sheet.ChartObjects().Height = sheet.Range("1:15").Height
}
```

```javascript
/*本示例设置图表工作表 Chart1 上所有图表的高度。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    chartobjects.Height = 350
}
```
