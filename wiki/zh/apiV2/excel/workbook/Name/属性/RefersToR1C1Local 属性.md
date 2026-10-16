---
title: "RefersToR1C1Local (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/RefersToR1C1Local%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DdfsD82c
---

# RefersToR1C1Local (属性)

返回或设置指定名称所引用的公式。公式以等号开头，由用户语言和 R1C1-样式引用组成。可读/写 **String** 类型。

## 示例

```javascript
/*本示例新建一个工作表，并将当前工作簿所有名称的列表插入到新工作表中，名称列表中包括其公式（由R1C1-样式引用和用户语言表示）。*/
function test() {
    let newSheet = ActiveWorkbook.Worksheets.Add()
    for (let i = 1; i <= ActiveWorkbook.Names.Count; i++) {
        newSheet.Cells.Item(i, 1).Value2 = ActiveWorkbook.Names.Item(i).NameLocal
        newSheet.Cells.Item(i, 2).Value2 = "'" + ActiveWorkbook.Names.Item(i).RefersToR1C1Local
    }
}
```

```javascript
/*本示例将索引号为2的名称所引用的公式设置为“=Sheet1!R11C4”（由R1C1-样式引用和用户语言表示）。*/
function test() {
    ActiveWorkbook.Names.Item(2).RefersToR1C1Local = "=Sheet1!R11C4"
}
```
