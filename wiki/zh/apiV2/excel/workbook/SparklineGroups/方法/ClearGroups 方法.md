---
title: "ClearGroups (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E6%96%B9%E6%B3%95/ClearGroups%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CTcOHOam
---

# ClearGroups (方法)

清除选定的迷你图组。

## 返回值

void

## 示例

```javascript
/*此示例清除活动单元格上迷你图组。*/
function test() {
    ActiveCell.SparklineGroups.ClearGroups()
}
```

```javascript
/*此示例清除活动工作表上单元格 H1 中迷你图组，并显示该区域上迷你图组的数量。*/
function test() {
    let sparklineGroups = ActiveSheet.Range("H1").SparklineGroups
    sparklineGroups.ClearGroups()
    console.log(sparklineGroups.Count)
}
```
