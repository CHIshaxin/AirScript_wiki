---
title: "Locked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
---
# Locked (属性)

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

## 说明

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中图表 Chart 2 是否已被锁定。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    console.log(chartobject.Locked)
}
```

```javascript
/*此示例取消图表工作表 Chart1 中图表的锁定。*/
function test() {
    let chartobject = Application.Charts.Item("Chart1").ChartObjects(1)
    chartobject.Locked = false
}
```
