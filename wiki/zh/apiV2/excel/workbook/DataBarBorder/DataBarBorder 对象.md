---
title: "DataBarBorder (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataBarBorder/DataBarBorder%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: ilu5eNpa
---

# DataBarBorder (对象)

表示由条件格式规则指定的数据条的边框。

## 说明

使用 **[DataBarBorder](./../DataBarBorder/DataBarBorder 对象.html)** 对象可获取或设置数据条的颜色和边框类型。若要访问与数据条条件格式规则关联的 **[DataBarBorder](./../DataBarBorder/DataBarBorder 对象.html)** 对象，请使用 **BarBorder** 属性。在检索 **[DataBarBorder](./../DataBarBorder/DataBarBorder 对象.html)** 对象后，使用其 **Color** 属性返回可用来设置数据条颜色的 **[FormatColor](./../FormatColor/FormatColor 对象.html)** 对象。

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
/*本示例判断如果活动工作表上区域 A1:A10 中第一个条件格式（数据条）的边框为实心边框，则设置该边框的颜色。*/
function test() {
    let dataBarBorder = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).BarBorder
    if (dataBarBorder.Type == xlDataBarBorderSolid) {
        dataBarBorder.Color.ColorIndex = 4
    }
}
```
