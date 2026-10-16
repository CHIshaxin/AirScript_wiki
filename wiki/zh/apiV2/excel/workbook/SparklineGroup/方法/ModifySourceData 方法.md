---
title: "ModifySourceData (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E6%96%B9%E6%B3%95/ModifySourceData%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D67RdgC9
---

# ModifySourceData (方法)

设置代表迷你图组源数据的区域。

## 返回值

void

SourceData - 代表源数据的区域。

## 示例

```javascript
/*此示例先选择一个位于 A1:A4 的迷你图组，然后使用区域 B1:D4 中的数据修改源数据，以使之增加一列。*/
function test() {
    Range("A1:A4").Select()
    ActiveCell.SparklineGroups.Item(1).ModifySourceData("B1:D4")
}
```

```javascript
/*本示例设置活动工作表上单元格 A1 中第一个迷你图组的源数据区域。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("A1").SparklineGroups.Item(1)
    sparklineGroup.ModifySourceData("B1:C4")
}
```
