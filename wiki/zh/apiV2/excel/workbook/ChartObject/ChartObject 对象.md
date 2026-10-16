---
title: "ChartObject (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/ChartObject%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: IzAvr55e
---

# ChartObject (对象)

代表工作表上的嵌入图表。

## 说明

**[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象充当 **[Chart](./../Chart/Chart 对象.html)** 对象的容器。 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象的属性和方法控制工作表上嵌入图表的外观和大小。 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象是 **[ChartObjects](./../ChartObjects/ChartObjects 对象.html)** 集合的成员。 **[ChartObjects](./../ChartObjects/ChartObjects 对象.html)** 集合包含单一工作表上的所有嵌入图表。

使用 **ChartObjects**(*index*)（其中 *index* 是嵌入图表的索引号或名称）可以返回单个 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象。

## 示例

```javascript
/*本示例将工作表 Sheet1 第一个图表的填充方案设置为浅色下对角线。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Chart.ChartArea.Format.Fill.Patterned(msoPatternLightDownwardDiagonal)
}
```

当选定嵌入图表时，其名称显示在“名称”框中。使用 **Name** 属性可设置或返回 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象的名称。

```javascript
/*此示例为工作表 Sheet1 的第二个图表添加圆角。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(2)
    chartobject.RoundedCorners = true
}
```
