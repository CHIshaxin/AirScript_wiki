---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataBarBorder/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Clcyc-T1
---

# Type (属性)

返回或设置由条件格式规则指定的数据条的边框类型。可读/写。

## 返回值

**[XlDataBarBorderType](./../../枚举/XlDataBarBorderType 枚举.html)**

## 示例

```javascript
/*本示例选择一个单元格区域，将数据条条件格式规则添加到该区域，使用BarBorder属性检索与该规则关联的DataBarBorder对象，然后设置数据条的颜色、淡色和类型。*/
function test() {
    Range("B1:B10").Select()
    Range("B1:B10").Activate()
    let dataBar = Selection.FormatConditions.AddDatabar()
    let dataBarBorder = dataBar.BarBorder
    dataBarBorder.Type = xlDataBarBorderSolid
    dataBarBorder.Color.ThemeColor = xlThemeColorAccent2
    dataBarBorder.Color.TintAndShade = 0
}
```

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式（数据条）的边框的类型是否为xlDataBarBorderSolid。*/
function test() {
    let dataBarBorder = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).BarBorder
    console.log(dataBarBorder.Type == xlDataBarBorderSolid)
}
```
