---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除工作表 Sheet1 中第一张图表的模拟运算表。*/
function test() {
    let datatable = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.DataTable
    datatable.Delete()
}
```

```javascript
/*本示例判断如果图表工作表 Chart1 中图表的模拟运算表边框的颜色为红色，则删除该运算表。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    if (datatable.Border.ColorIndex == 3) {
        datatable.Delete()
    }
}
```
