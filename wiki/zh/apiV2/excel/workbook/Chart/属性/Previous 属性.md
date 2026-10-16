---
title: "Previous (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Previous%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CsHIYJ1m
---

# Previous (属性)

返回代表上一个工作表的 **[Worksheet](./../../Worksheet/Worksheet 对象.html)** 对象。

## 说明

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

## 示例

```javascript
/*本示例激活图表工作表 Chart1 上一个工作表。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Previous.Activate()
}
```

```javascript
/*本示例显示图表工作表 Chart1 上一个工作表是否为成绩表。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.Previous.Name == "成绩表")
}
```
