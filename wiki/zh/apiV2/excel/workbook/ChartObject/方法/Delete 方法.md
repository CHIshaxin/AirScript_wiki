---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CYMzHFwV
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除工作表 Sheet1 上嵌入的第一个图表。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Delete()
}
```

```javascript
/*本示例删除图表工作表 Chart1 上的图表 Chart 1 。*/
function test() {
    let chartobject = Application.Charts.Item("Chart1").ChartObjects("Chart 1")
    chartobject.Delete()
}
```
