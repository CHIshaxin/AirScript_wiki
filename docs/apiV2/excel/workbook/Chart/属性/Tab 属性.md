---
title: "Tab 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Tab%20%E5%B1%9E%E6%80%A7.html
---
# Tab (属性)

为图表返回一个 **Tab** 对象。

## 示例

```javascript
/*本示例将图表工作表 Chart1 选项的颜色索引设置为无颜色。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Tab.ColorIndex = xlColorIndexNone
}
```

```javascript
/*本示例将图表工作表 Chart1 的配色方案中主题颜色的枚举值设置为 xlThemeColorAccent1 。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Tab.ThemeColor = xlThemeColorAccent1
}
```
