---
title: "Markers (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkPoints/%E5%B1%9E%E6%80%A7/Markers%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CZMEDE74
---

# Markers (属性)

返回一个 **[SparkColor](./../../SparkColor/SparkColor 对象.html)** 对象，该对象代表迷你图上正数据点的标记的颜色和可见性。只读

## 说明

**Markers** 属性对应于 **“显示”** 部分中 **“标记”** 复选框的设置，以及功能区 **“迷你图工具设计”** 选项卡上 **“样式”** 部分中 **“标记颜色”** 下拉列表中 **“标记”** 项的设置。

## 返回值

SparkColor

## 示例

```javascript
/*本示例将区域 A1:A4 上第一个迷你图组的迷你图的正数据点设置为不可见。*/
function test() {
    let sparkColor = Range("A1:A4").SparklineGroups.Item(1).Points.Markers
    sparkColor.Visible = false
}
```

```javascript
/*本示例判断如果活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的正数据点可见，则设置该数据点的颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Markers
    if (sparkColor.Visible) {
        sparkColor.Color.Color = RGB(0, 255, 255)
    }
}
```
