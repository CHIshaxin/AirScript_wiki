---
title: "Index (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Icon/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DWq4c1z7
---

# Index (属性)

返回一个 **Long** 值，该值指明 **[IconSet](./../../IconSet/IconSet 对象.html)** 对象中 **[Icon](./../../Icon/Icon 对象.html)** 对象的索引号。只读。

## 示例

```javascript
/*本示例显示活动工作表上区域 B1:B10 中第一个图标集条件格式的图标集中第一个图标的索引。*/
function test() {
    let icon = ActiveSheet.Range("B1:B10").FormatConditions.Item(1).IconSet.Item(1)
    console.log(icon.Index)
}
```

```javascript
/*本示例将第一张工作表上区域 B1:B10 中第二个图标集条件格式的图标集中第二个图标的索引赋值给 D1 单元格。*/
function test() {
    let icon = Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(2).IconSet.Item(2)
    Range("D1").Value2 = icon.Index
}
```
