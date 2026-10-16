---
title: "Hyperlinks (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Hyperlinks%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CXT5XTID
---

# Hyperlinks (属性)

返回一个 **[Hyperlinks](./../../Hyperlinks/Hyperlinks 对象.html)** 集合，它代表图表的超链接。

## 示例

```javascript
/*本示例在图表工作表 Chart1 第一个形状中创建新的超链接。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.Hyperlinks.Add(chart.Shapes(1), "https://www.wps.cn/")
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个内嵌图表超链接的数量。*/
function test() {
    let chart = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart
    console.log(chart.Hyperlinks.Count)
}
```
