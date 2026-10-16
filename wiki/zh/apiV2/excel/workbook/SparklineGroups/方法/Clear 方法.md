---
title: "Clear (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: YvdliluP
---

# Clear (方法)

清除选定的迷你图。

## 返回值

void

## 示例

```javascript
/*此示例清除活动工作表上区域 A1:A2 上迷你图组的迷你图。*/
function test() {
    let sparklineGroups = ActiveSheet.Range("A1:A2").SparklineGroups
    sparklineGroups.Clear()
}
```

```javascript
/*此示例清除第一张工作表上区域 A1:A2 上迷你图组的迷你图，并显示该区域上迷你图组的数量。*/
function test() {
    let sparklineGroups = Worksheets.Item(1).Range("A1:A2").SparklineGroups
    sparklineGroups.Clear()
    console.log(sparklineGroups.Count)
}
```
