---
title: "ChartObjects 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/ChartObjects%20%E5%AF%B9%E8%B1%A1.html
---
# ChartObjects (对象)

由指定的图表工作表、对话框工作表或工作表上的所有 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象组成的集合。

## 说明

每个 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象都代表一个嵌入图表。 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象充当 **[Chart](./../Chart/Chart 对象.html)** 对象的容器。 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象的属性和方法控制工作表上嵌入图表的外观和大小。

使用 **ChartObjects** 方法返回 **[ChartObjects](./../ChartObjects/ChartObjects 对象.html)** 集合。

## 示例

```javascript
/*本示例删除工作表 Sheet1 中所有图表。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects().Delete()
}
```

不能使用 **[ChartObjects](./../ChartObjects/ChartObjects 对象.html)** 集合来调用以下属性和方法：

- **Locked** 属性
- **Placement** 属性
- **PrintObject** 属性

与早期版本不同， **[ChartObjects](./../ChartObjects/ChartObjects 对象.html)** 集合现在可以读取表示高度、宽度、左对齐和顶对齐的属性。

使用 **Add** 方法可创建一个新的空嵌入图表并将它添加到集合中。使用 **ChartWizard** 方法可添加数据并设置新图表的格式。

```javascript
/*本示例在工作表 Sheet1 中新建一个内嵌图表，然后以折线图形式添加单元格 A1:A20 中的数据。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects().Add(100, 30, 400, 250)
    chartobject.Chart.ChartWizard(Worksheets.Item("Sheet1").Range("A1:A20"), xlLine, null, null, null, null, null, "New Chart", null, null, null)
}
```

使用 **ChartObjects**(*index*)（其中 *index* 是嵌入图表的索引号或名称）可以返回单个对象。

```javascript
/*本示例将工作表 Sheet1 上第一张图表填充方案设置为浅色下对角线。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartArea.Format.Fill.Patterned(msoPatternLightDownwardDiagonal)
}
```
