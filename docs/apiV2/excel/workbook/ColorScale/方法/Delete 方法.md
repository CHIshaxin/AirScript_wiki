---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除指定的条件格式规则对象。

## 示例

```javascript
/*本示例删除第一张工作表上区域 A1:A10 中第一个色阶条件格式。*/
function test() {
    let colorScale = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    colorScale.Delete()
}
```

```javascript
/*本示例删除区域 A1:A10 中第二个色阶条件格式。*/
function test() {
    Range("A1:A10").FormatConditions.Item(2).Delete()
}
```
