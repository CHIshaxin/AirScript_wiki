---
title: "Range (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DdJPXn2Y
---

# Range (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表在以上列表中的指定列表对象所应用的区域。

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第三列单元格的宽度。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(3).Range.Width)
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第一列单元格边框的粗细设置为粗线。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(1).Range.Borders.Weight = xlThick
}
```
