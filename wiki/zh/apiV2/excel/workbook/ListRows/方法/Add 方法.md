---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRows/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CFSIMoFv
---

# Add (方法)

向指定的 **ListObject**所代表的表格中添加新行。

## 说明

如果不指定 *Position*，新行将添加到底部。如果不指定 *AlwaysInsert*，表格下面的单元格将下移一行（与指定 **True** 的作用相同）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Position | any | 可选 | Integer 类型。指定新行的相对位置。 |
| AlwaysInsert | any | 可选 | Boolean。指定在插入新行时，是否始终移动表格中最后一行下面的单元格中的数据，而不考虑表格下面的行是否为空行。如果为 True，则表格下面的单元格将下移一行。如果为 False，并且表格下面的行为空行，则表格将扩展以占用该空行而不移动它下面的单元格；但如果表格下面的行包含数据，则在插入新行时，这些单元格将下移。 |

## 返回值

一个代表新行的 ListRow 对象。

## 示例

```javascript
/*以下示例向工作簿的第一张工作表中的默认 ListObject 对象添加新行。由于未指定位置，新行将添加到列表的底部。*/
function test() {
    let row = Application.ActiveWorkbook.Worksheets.Item(1).ListObjects.Item(1).ListRows.Add()
}
```

```javascript
/*下例给活动工作表的第一个 ListObject 对象添加一个新行，并加到列表中现有行的第四个位置。*/
function test() {
    ActiveSheet.ListObjects.Item(1).ListRows.Add(4, false)
}
```
