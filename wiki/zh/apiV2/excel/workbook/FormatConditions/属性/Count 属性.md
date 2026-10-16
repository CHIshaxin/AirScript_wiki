---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: qi0AOHT6
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示活动工作表上单元格区域 A1:A10 条件格式的数量。*/
function test() {
    console.log(ActiveSheet.Range("A1:A10").FormatConditions.Count)
}
```
