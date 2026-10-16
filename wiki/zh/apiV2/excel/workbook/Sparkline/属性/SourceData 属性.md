---
title: "SourceData (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sparkline/%E5%B1%9E%E6%80%A7/SourceData%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DgKMjXkd
---

# SourceData (属性)

返回或设置包含单个迷你图的源数据的区域。可读/写。

## 说明

单个迷你图的数据源区域必须是连续的。

## 返回值

**String**

## 示例

```javascript
/*本示例显示单元格 I1 中第一个迷你图组的第一个迷你图的源数据的区域。*/
function test() {
    console.log(Range("I1").SparklineGroups(1).Item(1).SourceData)
}
```

```javascript
/*本示例设置活动工作表上单元格 I1 中第一个迷你图组的第一个迷你图的源数据的区域。*/
function test() {
    let sparkline = ActiveSheet.Range("I1").SparklineGroups(1).Item(1)
    sparkline.SourceData = "A1:C1"
}
```
