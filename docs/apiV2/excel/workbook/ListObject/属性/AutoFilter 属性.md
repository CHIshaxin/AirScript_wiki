---
title: "AutoFilter 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/AutoFilter%20%E5%B1%9E%E6%80%A7.html
---
# AutoFilter (属性)

使用“自动筛选”筛选一个列表。只读。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张列表的自动筛选范围中的筛选器的数量。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.AutoFilter.Filters.Count)
}
```

```javascript
/*本示例选中活动工作表上包含 A2 单元格的列表的自动筛选应用到的范围。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A2").ListObject
    listObj.AutoFilter.Range.Select()
}
```
