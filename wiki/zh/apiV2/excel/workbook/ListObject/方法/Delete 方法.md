---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: JcZ6F2aQ
---

# Delete (方法)

删除 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象并从工作表中清除单元格数据。

## 说明

如果列表链接到 SharePoint 网站，删除列表不会影响运行 SharePoint Foundation 的服务器上的数据。对本地列表所做的任何未提交的更改不会发送到 SharePoint 列表（系统不会警告这些未提交的更改将会丢失）。

## 示例

```javascript
/*本示例删除工作表 Sheet1 上第一张列表。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.Delete()
}
```

```javascript
/*本示例判断如果活动工作表上包含 A1 单元格的列表的名称为“表1”，则删除该列表。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A1").ListObject
    if (listObj.Name == "表1") {
        listObj.Delete()
    }
}
```
