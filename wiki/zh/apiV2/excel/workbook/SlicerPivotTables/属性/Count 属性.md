---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerPivotTables/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bdbd7R1_
---

# Count (属性)

返回集合中对象的数目。

## 返回值

Long

## 示例

```javascript
/*本示例显示活动工作表上数据透视表的第一个切片器的缓存关联的数据透视表的数量。*/
function test() {
    let sliCache = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache
    console.log(sliCache.PivotTables.Count)
}
```
