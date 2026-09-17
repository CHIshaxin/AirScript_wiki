---
title: "DataFields 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/DataFields%20%E6%96%B9%E6%B3%95.html
---
# DataFields (方法)

返回一个对象，该对象表示单个数据透视表字段（ **[PivotField](./../../PivotField/PivotField 对象.html)** 对象）或当前显示为数据字段的所有字段集合（ **[PivotFields](./../../PivotFields/PivotFields 对象.html)** 对象）。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 字段的名称或编号（可为指定多个字段的数组）。 |

## 示例

```javascript
/*本示例向一张新工作表中添加数据透视表数据字段的名称列表。*/
function test() {
    let nwSheet = Worksheets.Add()
    nwSheet.Activate()
    let pvtTable = Worksheets.Item("Sheet1").Range("A4").PivotTable
    let rw = 0
    for (let i = 1; i <= pvtTable.DataFields().Count; i++) {
        rw++
        nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.DataFields(i).Name
    }
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表数据字段的数量*/
function test() {
    console.log(ActiveSheet.PivotTables(1).DataFields().Count)
}
```
