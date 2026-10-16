---
title: "Shadow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 0onbXMO2
---

# Shadow (属性)

返回或设置 **Boolean** 值，它确定字体是否是阴影字体或对象是否带有阴影。

## 说明

对于 **[Font](./../../Font/Font 对象.html)** 对象，该属性在 Windows 中无效，但保留其值（可被设置或返回）。

## 示例

```javascript
/*此示例将工作表 Sheet1 上图表 Chart 2 的阴影关闭。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    chartobject.Shadow = false
}
```

```javascript
/*此示例将工作表 Sheet1 上第一个图表设置为有阴影，并设置阴影样式。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Shadow = true
    chartobject.Shadow.Type = msoShadow8
}
```
