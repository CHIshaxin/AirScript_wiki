---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axes/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 2sRB0fHh
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的坐标轴的数量。*/
function test() {
    let axes = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes()
    console.log(`坐标轴的数量：${axes.Count}`)
}
```
