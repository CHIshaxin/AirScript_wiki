---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 说明

**[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象包含图表区的线条、填充、效果和文本格式。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 中图表的标题前景填充色为绿色。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    charttitle.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
}
```

```javascript
/*本示例设置图表工作表 Chart2 中图表的标题边框线条的粗细。*/
function test() {
    let charttitle = Application.Charts.Item("Chart2").ChartObjects(1).Chart.ChartTitle
    charttitle.Format.Line.Weight = 2
}
```
