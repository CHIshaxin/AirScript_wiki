---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: IjJ8ulqQ
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的第三个数据序列的数据标签的数量。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(3).DataLabels()
    console.log(`数据标签的数量：${datalabels.Count}`)
}
```
