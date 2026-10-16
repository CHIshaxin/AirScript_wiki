---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回 **[SlicerCaches](./../../SlicerCaches/SlicerCaches 对象.html)** 集合中指定 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象的索引。只读

## 示例

```javascript
/*本示例显示活动工作簿“切片器_name”切片器缓存的索引号。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").Index)
}
```

```javascript
/*本示例判断如果第一张工作簿中切片器缓存“切片器_date”的索引号为2，则清除该切片器缓存的筛选器。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches("切片器_date")
    if (slcCache.Index == 2) {
        slcCache.ClearManualFilter()
    }
}
```
