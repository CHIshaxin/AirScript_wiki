---
title: "ZOrder (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/ZOrder%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BLRFO4fR
---

# ZOrder (属性)

返回指定对象的 z-次序位置。**Long** 型，只读。

## 说明

在任何对象集合中，z-次序尾端的对象为 *collection*(1)，z-次序前端的对象为 *collection*(*collection*.**Count**)。例如，如果活动工作表中有嵌入图表，z-次序尾端的图表为 `ActiveSheet.ChartObjects(1)`，z-次序前端的图表为 `ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count)`。

## 示例

```javascript
/*此示例显示 Sheet1 上嵌入的第一张图表的 z-次序位置。*/
function test()
{
    console.log("The chart's z-order position is " +
    Application.Worksheets.Item("Sheet1").ChartObjects(1).ZOrder)
}
```
