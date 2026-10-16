---
title: "PivotRowAxis (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/PivotRowAxis%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B7vG8oiW
---

# PivotRowAxis (属性)

返回 **[PivotAxis](./../../PivotAxis/PivotAxis 对象.html)** 对象，该对象表示整个行轴。只读 **[PivotAxis](./../../PivotAxis/PivotAxis 对象.html)** 类型。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表中行轴上行的数量。*/
function test() {
    let ptb = Worksheets.Item("Sheet1").PivotTables(1)
    console.log(ptb.PivotRowAxis.PivotLines.Count)
}
```

```javascript
/*本示例显示活动工作表中名为“数据透视表1”的数据透视表中行轴上第一行的数据透视线的类型是否为“xlPivotLineRegular”。*/
function test() {
    let ptb = Application.ActiveSheet.PivotTables("数据透视表1")
    console.log(ptb.PivotRowAxis.PivotLines.Item(1).LineType == xlPivotLineRegular)
}
```
