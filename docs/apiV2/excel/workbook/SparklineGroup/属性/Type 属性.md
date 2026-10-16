---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

获取或设置组的迷你图的类型。可读/写。

## 返回值

XlSparkType

## 示例

```javascript
/*本示例设置活动工作表上单元格 J1 中第一个迷你图组的迷你图的类型。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("J1").SparklineGroups.Item(1)
    sparklineGroup.Type = xlSparkColumn
}
```

```javascript
/*本示例显示活动工作表上单元格 J1 中第一个迷你图组的迷你图的类型是否为折线图迷你图。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("J1").SparklineGroups.Item(1)
    console.log(sparklineGroup.Type == xlSparkLine)
}
```
