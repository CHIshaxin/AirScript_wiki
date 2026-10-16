---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

## 示例

```javascript
/*本示例判断如果第一张工作表上图表 Chart 1 的索引为 1 ，则显示其名称。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects("Chart 1")
    if (chartobject.Index == 1) {
        console.log(chartobject.Name)
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 上嵌入的图表 Chart 2 的索引是否为 2 。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    console.log(chartobject.Index == 2)
}
```
