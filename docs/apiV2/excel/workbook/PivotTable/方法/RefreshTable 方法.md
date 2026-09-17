---
title: "RefreshTable 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/RefreshTable%20%E6%96%B9%E6%B3%95.html
---
# RefreshTable (方法)

用源数据刷新数据透视表。如果成功，则返回 **True**。

## 返回值

Boolean

## 示例

```javascript
/*本示例刷新数据透视表。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("F4").PivotTable
    pvtTable.RefreshTable()
}
```

```javascript
/*本示例刷新活动工作表中每一张数据透视表，并显示是否成功*/
function test() {
    for (i = 1; i <= ActiveSheet.PivotTables().Count; i++) {
        if (ActiveSheet.PivotTables(i).RefreshTable()) {
            console.log(`第${i}张数据透视表刷新成功`)
        }
    }
}
```
