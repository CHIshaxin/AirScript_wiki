---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 说明

此属性对图表对象（嵌入式图表）而言是只读的。

## 示例

```javascript
/*此示例显示第一张工作表中第一张图表的名称。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects(1)
    console.log(chartobject.Name)
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的名称是否为 Chart 1 。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    console.log(chartobject.Name == "Chart 1")
}
```
