---
title: "Resize 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E6%96%B9%E6%B3%95/Resize%20%E6%96%B9%E6%B3%95.html
---
# Resize (方法)

**Resize** 方法允许 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象在新区域内调整大小。不插入或移动单元格。

## 说明

对于链接到正在运行 SharePoint Foundation 的服务器的表格，可使用此方法调整列表大小，方法是提供一个仅在所包含的行数方面与 **ListObject**的当前区域不同的 *Range* 参数。如果通过添加或删除列（在 *Range* 参数中）来调整链接到 SharePoint Foundation 的列表的大小，将导致运行时错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Range | Range | 必填 | 新区域。 |

## 示例

```javascript
/*本示例使用了 Resize 方法来调整活动工作簿的 Sheet1 中默认 ListObject 对象的大小。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.Resize(Range("A1:B10"))
}
```

```javascript
/*本示例为活动工作表上包含 A1 单元格的列表调整新范围。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A1").ListObject
    listObj.Resize(ActiveSheet.UsedRange)
}
```
