---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BMzvVMgC
---

# Delete (方法)

删除迷你图组。

## 返回值

void

## 示例

```javascript
/*本示例删除单元格 I1 中第一个迷你图组。*/
function test() {
    Range("I1").SparklineGroups.Item(1).Delete()
}
```

```javascript
/*本示例删除活动工作表上单元格 I1 中第一个迷你图组，并显示该迷你图组中迷你图的数量。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    sparklineGroup.Delete()
    console.log(sparklineGroup.Count)
}
```
