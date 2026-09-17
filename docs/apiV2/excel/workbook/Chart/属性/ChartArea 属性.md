---
title: "ChartArea 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ChartArea%20%E5%B1%9E%E6%80%A7.html
---
# ChartArea (属性)

返回一个 **[ChartArea](./../../ChartArea/ChartArea 对象.html)** 对象，该对象表示图表的整个图表区。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 的图表区域内部颜色设置为红色，并将其边框颜色设置为蓝色。*/
function test() {
    let area = Application.Charts.Item("Chart1").ChartArea
    area.Interior.ColorIndex = 3
    area.Border.ColorIndex = 5
}
```

```javascript
/*本示例显示图表工作表 Chart1 的图表区域是否有阴影。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartArea.Shadow)
}
```
