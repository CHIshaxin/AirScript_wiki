---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除指定的条件格式规则对象。

## 示例

```javascript
/*本示例删除工作表 Sheet1 上区域 B1:B10 中第一个（AboveAverage）条件格式。*/
function test() {
    let aboveAverage = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(1)
    aboveAverage.Delete()
}
```

```javascript
/*本示例判断如果活动工作表上区域 B1:B10 中第二个（AboveAverage）条件格式的类型为xlAboveAverageCondition，则删除该条件格式。*/
function test() {
    let aboveAverage = ActiveSheet.Range("B1:B10").FormatConditions.Item(2)
    if (aboveAverage.Type == xlAboveAverageCondition) {
        aboveAverage.Delete()
    }
}
```
