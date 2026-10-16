---
title: "SourceData (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/SourceData%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BiALo_uG
---

# SourceData (属性)

返回或设置包含迷你图组的源数据的区域。可读/写。

## 说明

源数据区域中的行数或列数必须与 **Location** 属性区域中的单元格数相等。

迷你图组内单个迷你图的数据源区域必须是连续的。

## 返回值

String

## 示例

```javascript
/*本示例显示包含活动工作表上单元格 J1 中第一个迷你图组的源数据的区域。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("J1").SparklineGroups.Item(1)
    console.log(sparklineGroup.SourceData)
}
```

```javascript
/*本示例设置包含活动工作表上单元格 J1 中第一个迷你图组的源数据的区域。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("J1").SparklineGroups.Item(1)
    sparklineGroup.SourceData = "A1:C1"
}
```
