---
title: "Characters 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E6%96%B9%E6%B3%95/Characters%20%E6%96%B9%E6%B3%95.html
---
# Characters (方法)

返回 **[Characters](./../../Characters/Characters 对象.html)** 对象，它代表对象文本内某个区域的字符。使用 **[Characters](./../../Characters/Characters 对象.html)** 对象可为文本字符串内的字符设置格式。

## 说明

**[Characters](./../../Characters/Characters 对象.html)** 对象不是集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Start | any | 可选 | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。 |
| Length | any | 可选 | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（*Start* 字符之后的所有字符）。 |

## 示例

```javascript
/*下例将 Sheet1 上的第一个图表数值轴上的单位标签中部分文本加粗。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Characters(2, 2).Font.Bold = true
}
```

```javascript
/*本示例显示图表工作表 Chart1 中数值轴上的单位标签前三个字符。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Characters(null, 3).Text)
}
```
