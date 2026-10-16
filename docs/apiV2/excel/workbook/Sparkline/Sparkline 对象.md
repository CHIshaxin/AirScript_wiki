---
title: "Sparkline 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sparkline/Sparkline%20%E5%AF%B9%E8%B1%A1.html
---
# Sparkline (对象)

表示单个迷你图。

## 说明

使用 **ModifyLocation** 方法可更改单个迷你图的位置，使用 **ModifySourceData** 方法可更改源数据区域。若要一次操作一组迷你图，请使用 **[SparklineGroup](./../SparklineGroup/SparklineGroup 对象.html)** 对象的成员。

## 示例

```javascript
/*本示例修改单元格 I1 中第一个迷你图组的第一个迷你图的位置，并显示修改后迷你图的位置的地址。*/
function test() {
    let sparkline = Range("I1").SparklineGroups(1).Item(1)
    sparkline.ModifyLocation(Range("K1"))
    console.log(sparkline.Location.Address())
}
```

```javascript
/*本示例修改单元格 I1 中第一个迷你图组的第一个迷你图的源数据，并显示修改后的源数据的区域。*/
function test() {
    let sparkline = Range("I1").SparklineGroups(1).Item(1)
    sparkline.ModifySourceData("Sheet1!A1:C1")
    console.log(`修改后的源数据区域：${sparkline.SourceData}`)
}
```
