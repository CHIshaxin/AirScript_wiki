---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BUsfN6No
---

# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除图表工作表 Chart1 的标题。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartTitle
    charttitle.Delete()
}
```

```javascript
/*本示例判断如果图表工作表 Chart1 的标题可见，则删除该图表的标题。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    if (chart.HasTitle) {
        chart.ChartTitle.Delete()
    }
}
```
