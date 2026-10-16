---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张图表的标题对象的名称。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    console.log(charttitle.Name)
}
```

```javascript
/*本示例显示第一张工作表上图表 Chart 2 的标题对象名称是否为标题。*/
function test() {
    let charttitle = Application.Sheets.Item(1).ChartObjects("Chart 2").Chart.ChartTitle
    console.log(charttitle.Name == "标题")
}
```
