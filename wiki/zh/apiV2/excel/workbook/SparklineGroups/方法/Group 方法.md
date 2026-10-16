---
title: "Group (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E6%96%B9%E6%B3%95/Group%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B4m8Y4ai
---

# Group (方法)

组合选定的迷你图。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Location | Range | 必填 | 组中第一个单元格的位置。 |

## 返回值

void

## 示例

```javascript
/*此示例选择区域 A1:A4 并组合该区域中的迷你图。*/
function test() {
    Range("A1:A4").Select()
    Selection.SparklineGroups.Group(Range("A1"))
}
```

```javascript
/*此示例组合活动工作表上区域 A1:A4 上迷你图组。*/
function test() {
    let sparklineGroups = ActiveSheet.Range("A1:A4").SparklineGroups
    sparklineGroups.Group(Range("A3:A4"))
}
```
