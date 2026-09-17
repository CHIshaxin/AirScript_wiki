---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheets/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示当前工作簿中工作表的数量*/
function test() {
    console.log("当前工作簿中工作表数量: " + Application.Worksheets.Count);
}
```
