---
title: "ModifyAppliesToRange (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E6%96%B9%E6%B3%95/ModifyAppliesToRange%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BR8E-p6q
---

# ModifyAppliesToRange (方法)

设置此格式规则所应用于的单元格区域。

## 说明

该区域必须采用 A1 引用样式，并全部包含在 **[FormatConditions](./../../FormatConditions/FormatConditions 对象.html)** 集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Range | Range | 必填 | 此格式规则将应用于的区域。 |

## 示例

```javascript
/*本示例设置第一张工作表上单元格区域 B1:B10 的第二个图标集条件格式所应用于的单元格区域。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(2)
    iconSet.ModifyAppliesToRange(Range("B1:B4"))
}
```

```javascript
/*本示例判断如果活动工作表上单元格区域 B1:B10 的第一个图标集条件格式使用的图标集的名称为xl3Flags，则修改该条件格式所应用于的单元格区域。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(1)
    if (iconSet.IconSet.ID == xl3Flags) {
        iconSet.ModifyAppliesToRange(Range("B7:B9"))
    }
}
```
