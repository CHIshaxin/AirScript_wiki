---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSets/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，该值指定可用于工作簿的图标集的数目。只读。

## 示例

```javascript
/*本示例显示活动工作簿上图标集的数目。*/
function test() {
    console.log(ActiveWorkbook.IconSets.Count)
}
```
