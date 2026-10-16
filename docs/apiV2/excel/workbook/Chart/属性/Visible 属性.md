---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 **[XlSheetVisibility](./../../枚举/XlSheetVisibility 枚举.html)** 值，它确定对象是否可见。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 为可见。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Visible = xlSheetVisible
}
```

```javascript
/*本示例显示图表工作表 Chart2 是否被隐藏 。*/
function test() {
    let chart = Application.Charts.Item("Chart2")
    console.log(chart.Visible == xlSheetHidden)
}
```
