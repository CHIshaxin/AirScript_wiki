---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObjects/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dz5kPfU5
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*本示例显示活动工作簿的 Sheet1 上的默认列表对象的名称。*/
function test() {
    let sheet = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    console.log(sheet.ListObjects.Item(1).Name)
}
```

```javascript
/*本示例为工作表 Sheet1 上第一张列表添加新列，并设置该列表的表样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ListColumns.Add()
    listObj.TableStyle = "TableStylePreset2_Accent1"
}
```
