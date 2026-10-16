---
title: "PageFields 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/PageFields%20%E6%96%B9%E6%B3%95.html
---
# PageFields (方法)

返回一个对象，该对象表示单个数据透视表字段（ **[PivotField](./../../PivotField/PivotField 对象.html)** 对象），或者当前显示为页字段的所有字段的一个集合（ **[PivotFields](./../../PivotFields/PivotFields 对象.html)** 对象）。只读。

## 说明

一个层次只能包含一个页字段。

对基于数据透视表高速缓存的数据透视表而言，返回的数据透视表字段的集合反应了当前高速缓存中的内容。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的字段的名称或编号（可以是指定多个字段的数组）。 |

## 示例

```javascript
/*本示例向一张新工作表上的列表中添加页字段名称。*/
function test() {
    let nwSheet = Application.Worksheets.Add()
    nwSheet.Activate()
    let pvtTable = Worksheets.Item("Sheet1").Range("F1").PivotTable
    let rw = 0
    for (let i = 1; i <= pvtTable.PageFields().Count; i++) {
        rw++
        nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PageFields(i).Name
    }
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的页面字段的数量*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PageFields().Count)
}
```
