---
title: "Cut (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E6%96%B9%E6%B3%95/Cut%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CXmL4bkJ
---

# Cut (方法)

将对象剪切到剪贴板。

## 说明

只可剪切嵌入式图表。

## 示例

```javascript
/*本示例将工作表 Sheet1 上所有图表剪切到剪贴板。*/
function test() {
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    chartobjects.Cut()
}
```

```javascript
/*本示例将图表工作表 Chart1 上所有图表剪贴到剪贴板。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    chartobjects.Cut()
}
```
