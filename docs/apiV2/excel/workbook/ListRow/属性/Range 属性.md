---
title: "Range 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRow/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
---
# Range (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表在以上列表中的指定列表对象所应用的区域。

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第四行单元格的高度。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListRows.Item(4).Range.Height)
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第一行单元格字体大小设为 25 磅。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.ListRows.Item(1).Range.Font.Size = 25
}
```
