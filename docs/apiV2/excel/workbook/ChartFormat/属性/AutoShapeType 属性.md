---
title: "AutoShapeType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/AutoShapeType%20%E5%B1%9E%E6%80%A7.html
---
# AutoShapeType (属性)

返回指定形状的类型。只读 **MsoAutoShapeType** 类型。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 上图表的第一个数据系列的第一个数据标签的形状是否为 msoShapeRectangle。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1).Format
    console.log(chartformat.AutoShapeType == msoShapeRectangle)
}
```

```javascript
/*此示例判断工作表 Sheet1 上第一张图表的第二个数据系列的第一个数据标签的形状是否为 msoShapeOval，并通知用户。*/
function test() {
    let chartformat = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1).Format
    if (chartformat.AutoShapeType == msoShapeOval) {
        console.log("数据标签的形状是 msoShapeOval")
    } else {
        console.log("数据标签的形状不是 msoShapeOval")
    }
}
```
