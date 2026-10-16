---
title: "Placement (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Placement%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 4rji_6aa
---

# Placement (属性)

返回或设置一个包含 **[XlPlacement](./../../枚举/XlPlacement 枚举.html)** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

## 示例

```javascript
/*本示例设置工作表 Sheet1 上所有图表大小固定，位置随单元格移动而变化。*/
function test() {
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    chartobjects.Placement = xlMove
}
```

```javascript
/*本示例显示第一张工作表上所有图表是否大小和位置随单元格而变化。*/
function test() {
    let chartobjects = Application.Worksheets.Item(1).ChartObjects()
    console.log(chartobjects.Placement == xlMoveAndSize)
}
```
