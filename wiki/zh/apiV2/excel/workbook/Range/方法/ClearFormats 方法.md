---
title: "ClearFormats (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: NtfNyQGW
---

# ClearFormats (方法)

清除对象的格式设置。

## 返回值

Variant

## 示例

```javascript
/*此示例清除工作表 Sheet1 上 A1:G37 单元格区域的所有格式设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:G37").ClearFormats()
}
```

```javascript
/*此示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects.Item(1).Chart.ChartArea.ClearFormats()
}
```
