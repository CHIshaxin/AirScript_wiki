---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconCriteria/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C-6HHZBV
---

# Count (属性)

返回一个用于指定图标集条件格式规则的条件数目的 **Long** 值。只读。

## 说明

条件数目等于为图标集分配的阈值的数目。

## 示例

```javascript
/*本示例显示第一张工作表上区域 A1:A10 中第一个图标集条件格式的条件数量。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    console.log(iconSet.IconCriteria.Count)
}
```
