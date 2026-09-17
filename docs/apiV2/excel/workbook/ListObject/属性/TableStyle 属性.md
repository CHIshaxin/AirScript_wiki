---
title: "TableStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/TableStyle%20%E5%B1%9E%E6%80%A7.html
---
# TableStyle (属性)

获取或设置指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的表样式。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例显示活动工作表上包含 B2 单元格的列表的表样式的名称。*/
function test() {
    let listObj = Application.ActiveSheet.Range("B2").ListObject
    console.log(listObj.TableStyle.Name)
}
```

```javascript
/*本示例设置工作表 Sheet1 上第一张列表的表样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.TableStyle = "TableStylePreset1_Accent1"
}
```
