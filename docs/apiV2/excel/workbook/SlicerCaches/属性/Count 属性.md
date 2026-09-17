---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCaches/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回集合中对象的数目。

## 返回值

Long

## 示例

```javascript
/*本示例显示活动工作簿上切片器缓存的数量。*/
function test() {
    console.log(`切片器缓存的数量：${ActiveWorkbook.SlicerCaches.Count}`)
}
```
