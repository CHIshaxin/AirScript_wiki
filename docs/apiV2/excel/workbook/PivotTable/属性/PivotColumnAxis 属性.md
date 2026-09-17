---
title: "PivotColumnAxis 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/PivotColumnAxis%20%E5%B1%9E%E6%80%A7.html
---
# PivotColumnAxis (属性)

返回 **[PivotAxis](./../../PivotAxis/PivotAxis 对象.html)** 对象，该对象表示整个列轴。只读 **[PivotAxis](./../../PivotAxis/PivotAxis 对象.html)** 类型。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表中列轴上行的数量。*/
function test() {
    let ptb = Worksheets.Item("Sheet1").PivotTables(1)
    console.log(ptb.PivotColumnAxis.PivotLines.Count)
}
```

```javascript
/*本示例显示活动工作表中名为“数据透视表1”的数据透视表中列轴上第一行的位置。*/
function test() {
    let ptb = Application.ActiveSheet.PivotTables("数据透视表1")
    console.log(ptb.PivotColumnAxis.PivotLines.Item(1).Position)
}
```
