---
title: "ClearContents 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ClearContents%20%E6%96%B9%E6%B3%95.html
---
# ClearContents (方法)

清除区域中的公式。

## 示例

```javascript
/*此示例清除工作表 Sheet1 上 A1:G37 单元格区域的公式，但保留其格式设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:G37").ClearContents()
}
```

```javascript
/*本示例将 A1 单元格的字体加粗，然后使用 ClearContents 清除 A1 单元格内容*/
function test() {
    Range("A1").Font.Bold = true
    Range("A1").ClearContents()
}
```
