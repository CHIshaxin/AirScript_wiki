---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItems/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回集合中对象的数量。只读。

## 返回值

Long

## 示例

```javascript
/*本示例显示第一张工作簿上名为“切片器_id”的切片器缓存中切片器项的数量。*/
function test() {
    let sliItems = Workbooks.Item(1).SlicerCaches("切片器_id").SlicerItems
    console.log(sliItems.Count)
}
```
