---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象样式。

## 返回值

Variant

## 示例

```javascript
/*本示例将B1单元格的样式删除。*/
function test() {
    Range("B1").Style.Delete()
}
```

```javascript
/*如果工作表 Sheet1 中的单元格 C1 当前应用的是“百分比”样式，则删除样式。*/
function test() {
    if (Worksheets.Item("Sheet1").Range("C1").Style == "Percent") {
        Worksheets.Item("Sheet1").Range("C1").Style.Delete()
    }
}
```
