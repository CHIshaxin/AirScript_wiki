---
title: "ShowAutoFilter (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowAutoFilter%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 719a0CoS
---

# ShowAutoFilter (属性)

返回一个 **Boolean** 类型的数值，该数值指示是否显示“自动筛选”。**Boolean** 类型，可读写。

## 说明

对于新的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象，**ShowAutoFilter** 属性默认为 **True**。

## 示例

```javascript
/*本示例显示活动工作簿的 Sheet1 中默认列表的 ShowAutoFilter 属性设置。*/
function test() {
    let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let oListCol = wrksht.ListObjects.Item(1)
    console.log(oListCol.ShowAutoFilter)
}
```

```javascript
/*本示例将活动工作表上包含 D2 单元格的列表设置为不展示“自动筛选”。*/
function test() {
    let listObj = Application.ActiveSheet.Range("D2").ListObject
    listObj.ShowAutoFilter = false
}
```
