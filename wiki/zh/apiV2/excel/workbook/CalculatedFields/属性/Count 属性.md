---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CalculatedFields/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BHwS-q5I
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*此示例显示活动工作表上的数据透视表计算字段的数量。*/
function test() {
    console.log(ActiveSheet.Range("I1").PivotTable.CalculatedFields().Count)
}
```
