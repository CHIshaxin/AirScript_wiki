---
title: "Shadow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BGCv1Xq3
---

# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*此示例判断如果图表工作表 Chart1 中图表的第一个数据序列的第一个数据标签的文本为“软件1班”，则将该数据标签设置为有阴影。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    if (datalabel.Caption == "软件1班") {
        datalabel.Shadow = true
    }
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个图表上的第三个系列的第一个数据标签是否有阴影。*/
function test() {
    let datalabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(3).DataLabels(1)
    console.log(datalabel.Shadow)
}
```
