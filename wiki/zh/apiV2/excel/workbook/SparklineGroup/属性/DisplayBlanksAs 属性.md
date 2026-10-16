---
title: "DisplayBlanksAs (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/DisplayBlanksAs%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: k2Hs0xtM
---

# DisplayBlanksAs (属性)

获取或设置空白单元格在迷你图组中的绘制方式。可以是 **XlDisplayBlanksAs** 常量之一。可读/写。

## 返回值

XlDisplayBlanksAs

## 示例

```javascript
/*本示例显示活动工作表上单元格 I1 中第一个迷你图组的DisplayBlanksAs属性是否为xlNotPlotted。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    console.log(sparklineGroup.DisplayBlanksAs == xlNotPlotted)
}
```

```javascript
/*本示例设置活动工作表上单元格 I1 中第一个迷你图组的DisplayBlanksAs属性。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    sparklineGroup.DisplayBlanksAs = xlInterpolated
}
```
