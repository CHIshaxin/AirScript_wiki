---
title: "PercentileValues (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E5%B1%9E%E6%80%A7/PercentileValues%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CPpTeWZu
---

# PercentileValues (属性)

返回或设置一个 **Boolean** 值，该值指明是否使用百分点确定图标集条件格式的阈值。

## 说明

如果未将每个阈值设置为百分点，则此属性将返回 **False**。如果将此属性设置为 **True**，则会将每个阈值类型设置为百分点。

## 示例

```javascript
/*本示例显示第一张工作表上单元格区域 B1:B10 的第一个图标集条件格式是否使用百分点确定阈值。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(1)
    console.log(iconSet.PercentileValues)
}
```

```javascript
/*本示例将活动工作表上单元格区域 B1:B10 的第二个图标集条件格式设置为使用百分点确定阈值。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(2)
    iconSet.PercentileValues = true
}
```
