---
title: "FilterCleared (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/FilterCleared%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DRLStCmH
---

# FilterCleared (属性)

返回是否清除切片器或时间线筛选器状态。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存是否清除切片器筛选器状态。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").FilterCleared)
}
```

```javascript
/*本示例显示第一张工作簿第三个切片器缓存是否清除切片器筛选器状态。*/
function test() {
    console.log(Workbooks.Item(1).SlicerCaches(3).FilterCleared)
}
```
