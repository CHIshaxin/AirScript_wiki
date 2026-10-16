---
title: "CustomDocumentProperties 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/CustomDocumentProperties%20%E5%B1%9E%E6%80%A7.html
---
# CustomDocumentProperties (属性)

返回或设置一个 **DocumentProperties** 集合，该集合表示指定工作簿的所有自定义文档属性。

## 说明

本属性返回的是自定义文档属性的整个集合。通过指定属性的名称或集合索引号（数字），用 **Item** 方法可返回集合中的单个成员（一个 **DocumentProperty** 对象）。

由于 **Item** 方法是 **DocumentProperties** 集合的默认方法，所以下列语句是相同的：

```javascript
CustomDocumentProperties.Item("Complete")
CustomDocumentProperties("Complete")
```

使用 **BuiltinDocumentProperties** 属性可返回内置文档属性的集合。

**msoPropertyTypeString** 类型的属性长度不能超过 255 个字符。

## 示例

```javascript
/*本示例在第一个工作表中列表显示自定义文档属性的名称和值。。*/
function test() {
    let rw = 1
    Worksheets.Item(1).Activate()
    for (i = 1; i <= ActiveWorkbook.CustomDocumentProperties.Count; i++) {
        let p = ActiveWorkbook.CustomDocumentProperties.Item(i)
        Cells(rw, 1).Value(undefined, p.Name)
        Cells(rw, 2).Value(undefined, p.Value)
        rw = rw + 1
    }
}
```

```javascript
/*本示例演示CustomDocumentProperties属性的总数，并向活动工作簿CustomDocumentProperties属性中添加新的自定义文档属性Complete，并告知用户。*/
function test() {
    console.log(ActiveWorkbook.CustomDocumentProperties.Count)
    let ret = ActiveWorkbook.CustomDocumentProperties.Add("Complete", false, undefined, 1)
    if (ret) {
        console.log("Complete自定义文档属性添加成功")
    }
}
```
