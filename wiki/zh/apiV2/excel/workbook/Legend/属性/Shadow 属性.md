---
title: "Shadow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CK7D1Qvv
---

# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*此示例向图表工作表 Chart1 的图例添加阴影。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Shadow = true
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的图例是否有阴影。*/
function test() {
    let legend = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.Shadow)
}
```
