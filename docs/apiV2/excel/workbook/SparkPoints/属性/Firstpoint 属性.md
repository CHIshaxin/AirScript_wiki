---
title: "Firstpoint 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkPoints/%E5%B1%9E%E6%80%A7/Firstpoint%20%E5%B1%9E%E6%80%A7.html
---
# Firstpoint (属性)

返回一个 **[SparkColor](./../../SparkColor/SparkColor 对象.html)** 对象，该对象代表迷你图上第一个数据点的标记的颜色和可见性。只读

## 说明

**Firstpoint** 属性对应于 **“显示”** 部分中 **“首点”** 复选框的设置，以及功能区 **“迷你图工具设计”** 选项卡上 **“样式”** 部分中 **“标记颜色”** 下拉列表中 **“首点”** 项的设置。

## 返回值

SparkColor

## 示例

```javascript
/*本示例显示区域 A1:A4 上第一个迷你图组的迷你图的第一个数据点是否可见。*/
function test() {
    let sparkColor = Range("A1:A4").SparklineGroups.Item(1).Points.Firstpoint
    console.log(sparkColor.Visible)
}
```

```javascript
/*本示例判断如果活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的第一个数据点可见，则设置该数据点的颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Firstpoint
    if (sparkColor.Visible) {
        sparkColor.Color.ColorIndex = 5
    }
}
```
