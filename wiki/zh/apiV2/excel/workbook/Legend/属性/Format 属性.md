---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CM_dcaH_
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将工作表一上嵌入图表一的图例的前景色设置为蓝色。*/
function test() {
    let format = Worksheets.Item(1).ChartObjects(1).Chart.Legend.Format
    format.Fill.ForeColor.RGB = RGB(0, 0, 255)
}
```

```javascript
/*此示例给图表工作表 Chart1 上的第一个图表中的图例添加阴影，并向右侧偏移 5 磅。*/
function test() {
    let format = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Format
    format.Shadow.OffsetX = 5
}
```
