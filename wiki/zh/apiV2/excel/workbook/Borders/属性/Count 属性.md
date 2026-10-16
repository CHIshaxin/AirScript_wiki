---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BUKI2_9m
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示活动工作簿的第一个样式对象的边框数量。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).Borders.Count)
}
```
