---
title: "ShowTotals (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowTotals%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 6SylJdK1
---

# ShowTotals (属性)

获取或设置一个 **Boolean** 类型的数值以指示“总计”行是否可见。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例显示了对活动工作簿的 Sheet1 中默认列表的 ShowTotals 属性的当前设置。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.ShowTotals)
}
```

```javascript
/*本示例将活动工作表上包含 C1 单元格的列表的“总计”行设置为可见。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C1").ListObject
    listObj.ShowTotals = true
}
```
