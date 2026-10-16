---
title: "LineWeight (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/LineWeight%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DhQuqpHT
---

# LineWeight (属性)

获取或设置迷你图组中迷你图的粗细。可读/写。

## 返回值

Variant

## 示例

```javascript
/*本示例显示活动工作表上单元格 I1 中第一个迷你图组的迷你图的粗细。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    console.log(sparklineGroup.LineWeight)
}
```

```javascript
/*本示例设置活动工作表上单元格 I1 中第一个迷你图组的迷你图的粗细。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    sparklineGroup.LineWeight = 2
}
```
