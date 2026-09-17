---
title: "DataBodyRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E5%B1%9E%E6%80%A7/DataBodyRange%20%E5%B1%9E%E6%80%A7.html
---
# DataBodyRange (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象为列的数据部分的大小。只读。

## 说明

返回的对象不包含标题单元格和总计单元格。

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第二列数据单元格的数量，不包含标题单元格和总计单元格。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(2).DataBodyRange.Count)
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第四列单元格字体大小设为 20 磅，不包含标题单元格和总计单元格。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(4).DataBodyRange.Font.Size = 20
}
```
