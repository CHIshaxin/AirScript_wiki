---
title: "InsertRowRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/InsertRowRange%20%E5%B1%9E%E6%80%A7.html
---
# InsertRowRange (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象代表指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的“插入”行（如果有）。**Range** 类型，只读。

## 说明

如果由于列表（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）处于非活动状态而使得“插入”行不可见，则返回 **null** 对象。

## 示例

```javascript
/*本示例激活了由活动工作簿第一个工作表内默认 ListObject 对象中 InsertRowRange 指定的区域。*/
function test() {
    let objList = Application.ActiveWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    let objListRng = objList.InsertRowRange
    if (objListRng === null) {
        return false
    } else {
        objListRng.Activate()
        return true
    }
}
```

```javascript
/*本示例为活动工作表上包含 A1 单元格的列表的插入行的第一个单元格赋值。*/
function test() {
    let objList = Application.ActiveSheet.Range("A1").ListObject
    objList.InsertRowRange.Cells(1, 1).Value2 = "100"
}
```
