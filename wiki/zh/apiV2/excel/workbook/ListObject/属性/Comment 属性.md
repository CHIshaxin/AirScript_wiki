---
title: "Comment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/Comment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Du2l6Sya
---

# Comment (属性)

返回或设置与列表对象关联的批注。可读/写 **String** 类型。

## 说明

批注文本不能超过 255 个字符。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张表格的批注。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.Comment)
}
```

```javascript
/*本示例设置工作表 Sheet1 上包含单元格 A2 的表格的批注，并显示该批注。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").Range("A2").ListObject
    listObj.Comment = "示例数据"
    console.log(listObj.Comment)
}
```
