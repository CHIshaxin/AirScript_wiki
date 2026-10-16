---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个 **[SparklineGroup](./../../SparklineGroup/SparklineGroup 对象.html)** 对象。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 指定元素在集合中的位置。 |

## 返回值

SparklineGroup

## 示例

```javascript
/*本示例删除活动工作表上单元格 H1 中第一个迷你图组，并显示单元格 H1 中迷你图组的数量。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("H1").SparklineGroups.Item(1)
    sparklineGroup.Delete()
    console.log(ActiveSheet.Range("H1").SparklineGroups.Count)
}
```

```javascript
/*本示例设置活动工作表上单元格 H1 中第一个迷你图组的主系列颜色。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("H1").SparklineGroups.Item(1)
    sparklineGroup.SeriesColor.Color = RGB(255, 255, 0)
}
```
