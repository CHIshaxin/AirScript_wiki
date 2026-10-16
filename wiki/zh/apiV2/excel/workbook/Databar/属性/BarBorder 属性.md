---
title: "BarBorder (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/BarBorder%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BJAO1R13
---

# BarBorder (属性)

返回指定数据条边框的对象。只读。

## 返回值

**[DataBarBorder](./../../DataBarBorder/DataBarBorder 对象.html)**

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
/*本示例设置活动工作表上区域 A1:A10 中第一个为数据条的条件格式的边框的颜色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.BarBorder.Color.ColorIndex = 3
}
```
