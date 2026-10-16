---
title: "PlotBy (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/PlotBy%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D4SsGyhi
---

# PlotBy (属性)

返回或设置当迷你图所基于的数据处于方形区域中时，如何绘制迷你图。可读/写。

## 说明

只有在迷你图所基于的数据处于方形区域中时，才可以设置此属性，例如，数据位于区域 A1:B2。

方形区域中数据的默认值为按行绘制数据 (**xlSparklineRowsSquare**)。

## 返回值

**[XlSparklineRowCol](./../../枚举/XlSparklineRowCol 枚举.html)**

## 示例

```javascript
/*本示例显示活动工作表上单元格 I1 中第一个迷你图组的PlotBy属性是否为xlSparklineNonSquare。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    console.log(sparklineGroup.PlotBy == xlSparklineNonSquare)
}
```

```javascript
/*本示例设置活动工作表上单元格区域 I1:J1 中第一个迷你图组的PlotBy属性。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1:J1").SparklineGroups.Item(1)
    sparklineGroup.PlotBy = xlSparklineRowsSquare
}
```
