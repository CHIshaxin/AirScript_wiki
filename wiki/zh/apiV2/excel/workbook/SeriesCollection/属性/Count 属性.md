---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesCollection/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C8tGGMhN
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*此示例显示第一张工作表中第一个图表的系列数量。*/
function test() {
    console.log(Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection().Count)
}
```
