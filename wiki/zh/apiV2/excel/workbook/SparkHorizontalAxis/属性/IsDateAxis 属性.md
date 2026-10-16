---
title: "IsDateAxis (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkHorizontalAxis/%E5%B1%9E%E6%80%A7/IsDateAxis%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BLG_GLYn
---

# IsDateAxis (属性)

返回迷你图的水平轴是否基于日期值。只读。

## 说明

如果水平轴是基于日期值，则为 **True**；否则为 **False**。

## 返回值

**Boolean**

## 示例

```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组的水平轴是否基于日期值。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    console.log(shAxis.IsDateAxis)
}
```

```javascript
/*本示例判断如果活动工作表上单元格 H1 中第一个迷你图组的水平轴可见，则显示该水平轴是否基于日期值。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    if (shAxis.Axis.Visible) {
        console.log(shAxis.IsDateAxis)
    }
}
```
