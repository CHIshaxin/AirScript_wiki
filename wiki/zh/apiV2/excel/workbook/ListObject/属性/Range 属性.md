---
title: "Range (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Du9V3CDY
---

# Range (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表在以上列表中的指定列表对象所应用的区域。

## 示例

```javascript
/*本示例显示活动工作表上包含单元格 C1 的列表应用到的区域的地址。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C1").ListObject
    console.log(listObj.Range.Address())
}
```

```javascript
/*本示例选中工作表 Sheet1 中第一张列表应用到的区域。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.Range.Select()
}
```
