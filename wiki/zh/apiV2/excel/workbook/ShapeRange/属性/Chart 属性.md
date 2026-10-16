---
title: "Chart (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Chart%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CxVm9tpz
---

# Chart (属性)

返回 **[Chart](./../../Chart/Chart 对象.html)** 对象，该对象代表形状区域中包含的图表。只读。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象的图表标题设置为“销售数量”。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1])
    let chart = shpRange.Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "销售数量"
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象的图表类别坐标轴标题设置为“月份”。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2])
    let chart = shpRange.Chart
    chart.Axes(xlCategory).HasTitle = true
    chart.Axes(xlCategory).AxisTitle.Text = "月份"
}
```
