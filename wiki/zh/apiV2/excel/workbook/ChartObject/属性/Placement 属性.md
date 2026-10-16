---
title: "Placement (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Placement%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BCipo7m8
---

# Placement (属性)

返回或设置一个包含 **[XlPlacement](./../../枚举/XlPlacement 枚举.html)** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

## 示例

```javascript
/*本示例设置工作表 Sheet1 上第一个图表为可自由浮动（既不随下方单元格移动，也不随其改变大小）。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Placement = xlFreeFloating
}
```

```javascript
/*本示例显示第一张工作表上图表 Chart 2 是否随单元格移动和调整大小而变化。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects("Chart 2")
    console.log(chartobject.Placement == xlMoveAndSize)
}
```
