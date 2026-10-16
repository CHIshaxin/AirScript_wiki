---
title: "Shadow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 0N0w0eAw
---

# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*此示例向图表工作表 Chart1 的标题添加阴影。*/
function test() {
    Application.Charts.Item("Chart1").ChartTitle.Shadow = true
}
```

```javascript
/*此示例向工作表 Sheet1 中第一张图表的标题添加阴影，并设置阴影格式。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    charttitle.Shadow = true
    charttitle.Shadow.Type = msoShadow8
}
```
