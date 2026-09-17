---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRows/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Integer** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示活动工作表上包含 C2 单元格的列表中行的数量。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C2").ListObject
    console.log(listObj.ListRows.Count)
}
```
