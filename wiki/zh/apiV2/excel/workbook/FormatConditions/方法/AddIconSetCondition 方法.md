---
title: "AddIconSetCondition (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/AddIconSetCondition%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B9mZ19HH
---

# AddIconSetCondition (方法)

返回一个代表指定区域图标集条件格式规则的新 **[IconSetCondition](./../../IconSetCondition/IconSetCondition 对象.html)** 对象。

## 说明

使用图标集为数据添加注释并将数据分为按阈值隔开的三到五类数据。每种图标均代表某一值范围。例如，在 3 箭头图标集中，红色向上箭头代表较高值，黄色侧箭头代表中间值，绿色向下箭头代表较低值。

## 示例

```javascript
/*本示例在活动工作表上单元格区域 A1:A10 上新增图标集条件格式。*/
function test() {
    ActiveSheet.Range("A1:A10").FormatConditions.AddIconSetCondition()
}
```

```javascript
/*本示例在活动工作表上单元格区域 C1:C10 上新增图标集条件格式，并设置该条件格式使用的图标集。*/
function test() {
    let iconSetCondition = ActiveSheet.Range("C1:C10").FormatConditions.AddIconSetCondition()
    iconSetCondition.IconSet = ActiveWorkbook.IconSets(xl4Arrows)
}
```
