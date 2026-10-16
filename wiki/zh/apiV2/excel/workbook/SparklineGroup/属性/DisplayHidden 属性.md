---
title: "DisplayHidden (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/DisplayHidden%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CxZhZD4d
---

# DisplayHidden (属性)

指定是否在迷你图组中绘制隐藏单元格。可读/写。

## 说明

如果隐藏单元格在迷你图组中显示，则为 **True**。

## 返回值

Boolean

## 示例

```javascript
/*本示例将活动工作表上单元格 I1 中第一个迷你图组设置为绘制隐藏单元格。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    sparklineGroup.DisplayHidden = true
}
```

```javascript
/*本示例显示单元格 I1 中第一个迷你图组是否绘制隐藏单元格。*/
function test() {
    console.log(Range("I1").SparklineGroups.Item(1).DisplayHidden)
}
```
