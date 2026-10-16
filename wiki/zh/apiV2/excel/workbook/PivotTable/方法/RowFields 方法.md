---
title: "RowFields (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/RowFields%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BynIOOB7
---

# RowFields (方法)

返回一个对象，该对象表示数据透视表中的单个字段（ **[PivotField](./../../PivotField/PivotField 对象.html)** 对象），或者当前未显示为行字段的所有字段的一个集合（ **[PivotFields](./../../PivotFields/PivotFields 对象.html)** 对象）。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的字段的名称或编号（可以是指定多个字段的数组）。 |

## 示例

```javascript
/*本示例在新工作表上的清单中添加数据透视表的行字段名称。*/
function test() {
    let nwSheet = Application.Worksheets.Add()
    nwSheet.Activate()
    let pvtTable = Worksheets("Sheet1").Range("F1").PivotTable
    let rw = 0
    for (let i = 1; i <= pvtTable.RowFields().Count; i++) {
        rw = rw + 1
        nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.RowFields(i).Name
    }
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的行字段数量*/
function test() {
    console.log(ActiveSheet.PivotTables(1).RowFields().Count)
}
```
