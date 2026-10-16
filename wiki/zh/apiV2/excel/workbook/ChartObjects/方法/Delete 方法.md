---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BT8xZ0Pe
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例判断如果工作表 Sheet1 上图表的数量大于等于 2 ，则将该工作表上所有图表删除。*/
function test() {
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    if (chartobjects.Count >= 2) {
        chartobjects.Delete()
    }
}
```

```javascript
/*本示例删除图表工作表 Chart1 上所有图表。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    chartobjects.Delete()
}
```
