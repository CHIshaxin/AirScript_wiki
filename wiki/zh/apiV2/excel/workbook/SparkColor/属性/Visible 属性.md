---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkColor/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: zZjeM_gf
---

# Visible (属性)

返回或设置点是否可见。可读/写

## 说明

**Visible** 属性的设置对应于功能区 **“迷你图工具设计”** 选项卡 **“显示”** 部分中的 **“高点”**、**“低点”**、**“负点”**、**“首点”**、**“尾点”** 和 **­“标记”** 复选框的设置。使用 **[SparkPoints](./../../SparkPoints/SparkPoints 对象.html)** 对象的对应属性可设置这些复选框的状态。

## 返回值

Boolean

## 示例

```javascript
/*本示例将区域 A1:A4 上第一个迷你图组的迷你图的正数据点设置为不可见。*/
function test() {
    let sparkColor = Range("A1:A4").SparklineGroups.Item(1).Points.Markers
    sparkColor.Visible = false
}
```

```javascript
/*本示例显示区域 A1:A4 上第一个迷你图组的迷你图的负数据点是否可见。*/
function test() {
    let sparkColor = Range("A1:A4").SparklineGroups.Item(1).Points.Negative
    console.log(sparkColor.Visible)
}
```
