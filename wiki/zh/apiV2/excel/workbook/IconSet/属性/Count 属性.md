---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSet/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: eqF-0p7-
---

# Count (属性)

返回一个 **Long** 值，它指定图标集中图标的数目。只读。

## 示例

```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第一个图标集条件格式的图标集中图标的数量。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(1).IconSet
    console.log(iconSet.Count)
}
```
