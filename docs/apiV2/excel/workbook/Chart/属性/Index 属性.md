---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的索引号。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.Index)
}
```

```javascript
/*本示例显示图表工作表 Chart2 的索引号是否为2。*/
function test() {
    let chart = Application.Charts.Item("Chart2")
    console.log(chart.Index == 2)
}
```
