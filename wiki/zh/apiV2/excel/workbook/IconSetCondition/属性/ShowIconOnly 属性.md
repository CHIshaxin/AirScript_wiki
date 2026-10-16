---
title: "ShowIconOnly (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E5%B1%9E%E6%80%A7/ShowIconOnly%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bggof2CB
---

# ShowIconOnly (属性)

返回或设置一个 **Boolean** 值，该值指示是否仅显示图标集条件格式的图标。

## 说明

此属性的默认值为 **False**。

## 示例

```javascript
/*本示例显示第一张工作表上单元格区域 C1:C10 的第一个图标集条件格式是否仅展示图标。*/
function test() {
    let iconSet = Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(1)
    console.log(iconSet.ShowIconOnly)
}
```

```javascript
/*本示例将活动工作表上单元格区域 C1:C10 的第二个图标集条件格式设置为仅展示图标。*/
function test() {
    let iconSet = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    iconSet.ShowIconOnly = true
}
```
