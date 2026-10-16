---
title: "RefersToLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/RefersToLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BD1do3DO
---

# RefersToLocal (属性)

返回或设置指定名称所引用的公式。公式以等号开头，由用户语言和 A1-样式引用组成。**String** 类型，可读写。

## 示例

```javascript
/*本示例新建一个工作表，并将当前工作簿所有名称的列表插入到新工作表中，名称列表中包括其公式（由 A1-样式引用和用户语言组成）。*/
function test() {
    let newSheet = ActiveWorkbook.Worksheets.Add()
    for (let i = 1; i <= ActiveWorkbook.Names.Count; i++) {
        newSheet.Cells.Item(i, 1).Value2 = ActiveWorkbook.Names.Item(i).NameLocal
        newSheet.Cells.Item(i, 2).Value2 = "'" + ActiveWorkbook.Names.Item(i).RefersToLocal
    }
}
```

```javascript
/*本示例将索引号为1的名称所引用的公式设置为“=Sheet1!$A$4”（由 A1-样式引用和用户语言组成）。*/
function test() {
    ActiveWorkbook.Names.Item(1).RefersToLocal = "=Sheet1!$A$4"
}
```
