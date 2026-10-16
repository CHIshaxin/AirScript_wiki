---
title: "RefersTo (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/RefersTo%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CfY3qnFM
---

# RefersTo (属性)

用宏语言以 A1 样式表示法返回或设置名称所引用的公式（以等号开头）。**String** 类型，可读写。

## 示例

```javascript
/*本示例创建活动工作簿中所有名称的列表，并用宏语言以 A1 样式表示法表示这些名称所引用的公式。该列表将出现在本示例新建的工作表中。*/
function test() {
    let newSheet = Application.Worksheets.Add()
    let i = 1
    for (let nm = 1; nm <= Application.ActiveWorkbook.Names.Count; nm++) {
        newSheet.Cells.Item(i, 1).Value2 = Application.ActiveWorkbook.Names.Item(nm).Name
        newSheet.Cells.Item(i, 2).Value2 = "'" + Application.ActiveWorkbook.Names.Item(nm).RefersTo
        i = i + 1
    }
    newSheet.Columns.Item("A:B").AutoFit()
}
```

```javascript
/*本示例用宏语言以 A1 样式表示法设置索引号为3的名称所引用的公式。*/
function test() {
    ActiveWorkbook.Names.Item(3).RefersTo = "=Sheet1!$H$4"
}
```
