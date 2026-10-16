---
title: "Ungroup (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E6%96%B9%E6%B3%95/Ungroup%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DZD5u1qF
---

# Ungroup (方法)

对选定迷你图组中的迷你图取消组合。

## 返回值

void

## 示例

```javascript
/*本示例选择区域 A1:A4 并对该区域中的迷你图取消组合。*/
function test() {
    Range("A1:A4").Select()
    Selection.SparklineGroups.Ungroup()
}
```

```javascript
/*此示例组合活动工作表上区域 A1:A4 上迷你图组，并取消对 A1 单元格中迷你图组的组合。*/
function test() {
    let sparklineGroups = ActiveSheet.Range("A1:A4").SparklineGroups
    sparklineGroups.Group(Range("A3"))
    Range("A1").SparklineGroups.Ungroup()
}
```
