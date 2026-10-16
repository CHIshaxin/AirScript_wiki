---
title: "Color 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataBarBorder/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
---
# Color (属性)

返回一个指定数据条（由条件格式规则指定的数据条）边框色的对象。只读。

## 返回值

**[FormatColor](./../../FormatColor/FormatColor 对象.html)**

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
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式（数据条）的边框的颜色。*/
function test() {
    let dataBarBorder = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).BarBorder
    dataBarBorder.Color.ColorIndex = 3
}
```
