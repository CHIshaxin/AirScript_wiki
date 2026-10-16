---
title: "CalculatedItems (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CalculatedItems/CalculatedItems%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: 2MKEsafr
---

# CalculatedItems (对象)

由 **[PivotItem](./../PivotItem/PivotItem 对象.html)** 对象组成的集合，这些对象代表指定数据透视表中的所有计算项。

## 示例

```javascript
/*本示例为工作表 Sheet1 上数据透视表的第四个字段添加计算项。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields(4)
    pvtField.CalculatedItems().Add("公式1", "='21'+1")
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的字段“name”的“公式2”计算项的位置。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").CalculatedItems().Item("公式2")
    console.log(pvtItem.Position)
}
```
