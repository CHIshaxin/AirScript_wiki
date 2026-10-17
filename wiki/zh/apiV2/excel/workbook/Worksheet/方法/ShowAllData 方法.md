---
title: "ShowAllData (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/ShowAllData%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C18GoxLo
---

# ShowAllData (方法)

使当前筛选列表的所有行均可见。如果正在使用自动筛选，则本方法将下拉列表框内容改为“（全部）”。

## 示例

```javascript
/*本示例使工作表 Sheet1 上所有的数据可见。该示例应在包含使用 AutoFilter 命令筛选的列表的工作表上运行。*/
function test() {
    Application.Worksheets.Item("Sheet1").ShowAllData()
}
```

```javascript
/*本示例使工作簿的每一张工作表上筛选列表的所有行均可见。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        Application.Sheets.Item(i).ShowAllData()
    }
}
```
