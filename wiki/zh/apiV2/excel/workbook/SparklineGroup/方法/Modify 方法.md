---
title: "Modify (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E6%96%B9%E6%B3%95/Modify%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B8T3A8Jp
---

# Modify (方法)

设置迷你图组的位置和源数据。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Location | Range | 必填 | 代表迷你图组所在位置的 **[Range](./../../Range/Range 对象.html)** 对象。 |
| SourceData | string | 必填 | 代表迷你图组的源数据的区域。 |

## 返回值

void

## 示例

```javascript
/*此示例先选择一个位于 A1:A4 的迷你图组，然后通过将迷你图组的位置更改为 A1:A3 来删除一行数据。还必须对数据源进行修改，以使之仅包括前三行数据。*/
function test() {
    Range("A1:A4").Select()
    ActiveCell.SparklineGroups.Item(1).Modify(Range("A1:A3"), "Sheet1!B1:D3")
}
```

```javascript
/*本示例设置活动工作表上单元格 A1 中第一个迷你图组的位置和源数据。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("A1").SparklineGroups.Item(1)
    sparklineGroup.Modify(Range("A6"), "Sheet1!B1:D1")
}
```
