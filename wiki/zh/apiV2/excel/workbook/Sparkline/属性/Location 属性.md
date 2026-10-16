---
title: "Location (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sparkline/%E5%B1%9E%E6%80%A7/Location%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BFweltsx
---

# Location (属性)

返回或设置单个迷你图的位置。可读/写。

## 返回值

Range

## 示例

```javascript
/*本示例显示单元格 I1 中第一个迷你图组的第一个迷你图的位置的地址。*/
function test() {
    console.log(Range("I1").SparklineGroups(1).Item(1).Location.Address())
}
```

```javascript
/*本示例选中活动工作表上单元格 I1 中第一个迷你图组的第一个迷你图的位置。*/
function test() {
    let sparkline = ActiveSheet.Range("I1").SparklineGroups(1).Item(1)
    sparkline.Location.Select()
}
```
