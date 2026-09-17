---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObjects/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Integer** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示第一张工作表中列表的数量。*/
function test() {
    console.log(Worksheets.Item(1).ListObjects.Count)
}
```
