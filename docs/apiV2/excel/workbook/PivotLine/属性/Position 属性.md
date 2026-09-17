---
title: "Position 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLine/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
---
# Position (属性)

返回 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 对象的位置。只读。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第一条数据透视线的位置。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(1)
    console.log(pvtLine.Position)
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表中行轴上第一条数据透视线的位置。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    console.log(pvtLine.Position)
}
```
