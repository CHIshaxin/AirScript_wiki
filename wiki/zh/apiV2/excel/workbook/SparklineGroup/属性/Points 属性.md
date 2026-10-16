---
title: "Points (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/Points%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cjm_dV2W
---

# Points (属性)

返回迷你图组的关联 **[SparkPoints](./../../SparkPoints/SparkPoints 对象.html)** 对象。只读。

## 返回值

SparkPoints

## 示例

```javascript
/*本示例将活动工作表上单元格 I2 中第一个迷你图组的迷你图的高点设置为可见。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I2").SparklineGroups.Item(1)
    sparklineGroup.Points.Highpoint.Visible = true
}
```

```javascript
/*本示例设置活动工作表上单元格 I2 中第一个迷你图组的迷你图上第一个数据点的颜色。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I2").SparklineGroups.Item(1)
    sparklineGroup.Points.Firstpoint.Color.ColorIndex = 4
}
```
