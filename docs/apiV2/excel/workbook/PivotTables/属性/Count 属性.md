---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTables/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的数量。*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables().Count)
}
```
