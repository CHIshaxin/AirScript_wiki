---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

创建一个新的迷你图组并返回一个 **[SparklineGroup](./../../SparklineGroup/SparklineGroup 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | XlSparkType | 必填 | 迷你图的类型。 |
| SourceData | string | 必填 | 代表要用于创建迷你图的区域。 |

## 返回值

SparklineGroup

## 示例

```javascript
/*此示例向区域 A1:A4 中添加一个迷你图组。该组中的迷你图为列迷你图并且绑定到区域 B1:E4 中的数据。*/
function test() {
    Range("A1:A4").SparklineGroups.Add(xlSparkColumn, "Sheet1!B1:E4")
}
```

```javascript
/*此示例向活动工作表上区域 A1:A2 中添加一个迷你图组，并显示该工作表上单元格 A1 中迷你图组的数量。*/
function test() {
    ActiveSheet.Range("A1:A2").SparklineGroups.Add(xlSparkLine, "B1:D2")
    console.log(ActiveSheet.Range("A1").SparklineGroups.Count)
}
```
