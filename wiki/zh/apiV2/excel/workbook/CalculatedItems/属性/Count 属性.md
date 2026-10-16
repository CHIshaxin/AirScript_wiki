---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CalculatedItems/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DsPanf_q
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示活动工作表上第一张数据透视表的字段“name”中计算项的数量。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("name")
    console.log(pvtField.CalculatedItems().Count)
}
```
