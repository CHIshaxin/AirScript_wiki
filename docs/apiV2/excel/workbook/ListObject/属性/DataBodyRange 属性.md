---
title: "DataBodyRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/DataBodyRange%20%E5%B1%9E%E6%80%A7.html
---
# DataBodyRange (属性)

返回 **[Range](./../../Range/Range 对象.html)** 对象，它代表表中值的范围，但不包括标题行。只读。

## 示例

```javascript
/*此示例激活工作表 Sheet1，并选定活动工作表上第一张列表中的活动数据区域。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    ActiveSheet.ListObjects.Item(1).DataBodyRange.Select()
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张列表的值范围的地址。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.DataBodyRange.Address())
}
```
