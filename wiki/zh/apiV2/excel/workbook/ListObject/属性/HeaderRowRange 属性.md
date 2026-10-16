---
title: "HeaderRowRange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/HeaderRowRange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dv4QClEQ
---

# HeaderRowRange (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示列表（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）标题行的区域。**Range** 类型，只读。

## 示例

```javascript
/*本示例激活了由活动工作簿上工作表 Sheet1 中默认 ListObject 对象的 HeaderRowRange 属性指定的区域。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    let objListRng = listObj.HeaderRowRange
    objListRng.Activate()
}
```

```javascript
/*本示例显示活动工作表上包含 A2 单元格的列表的标题行的值。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A2").ListObject
    console.log(listObj.HeaderRowRange.Value2.toString())
}
```
