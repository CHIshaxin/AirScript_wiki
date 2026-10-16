---
title: "ShowTableStyleFirstColumn (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowTableStyleFirstColumn%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: xuUpJlcL
---

# ShowTableStyleFirstColumn (属性)

返回或设置是否显示指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的第一列。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作表上包含 B1 单元格的列表是否展示第一列。*/
function test() {
    let listObj = Application.ActiveSheet.Range("B1").ListObject
    console.log(listObj.ShowTableStyleFirstColumn)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一张列表设置为展示第一列。*/
function test() {
    let listObj = Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ShowTableStyleFirstColumn = true
}
```
