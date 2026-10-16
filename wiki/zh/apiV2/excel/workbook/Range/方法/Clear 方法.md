---
title: "Clear (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D9YiRdvv
---

# Clear (方法)

清除整个对象。

## 示例

```javascript
/*此示例清除工作表 Sheet1 中 A1:G37 单元格区域的公式和格式设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:G37").Clear()
}
```

```javascript
/*此示例设置工作表 Sheet1 中 H2 单元格的值为 abc 并清除 H1 单元格所在列的内容。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("H2").Value2 = "abc"
    Application.Worksheets.Item("Sheet1").Range("H:H").Clear()
}
```
