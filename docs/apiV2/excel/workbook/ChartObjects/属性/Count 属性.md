---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中图表的数量。*/
function test() {
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    console.log(`工作表 Sheet1 中有 ${chartobjects.Count} 个图表`)
}
```

```javascript
/*本示例显示第一张工作表上图表的数量是否为 2 。*/
function test() {
    let chartobjects = Application.Sheets.Item(1).ChartObjects()
    console.log(chartobjects.Count == 2)
}
```
