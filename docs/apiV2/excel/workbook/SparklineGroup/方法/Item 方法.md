---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

返回一个 **[Sparkline](./../../Sparkline/Sparkline 对象.html)** 对象。只读

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 迷你图的索引值。 |

## 返回值

Sparkline

## 示例

```javascript
/*本示例显示单元格 I1 中第一个迷你图组的第一个迷你图的位置的地址。*/
function test() {
    console.log(Range("I1").SparklineGroups.Item(1).Item(1).Location.Address())
}
```

```javascript
/*本示例设置活动工作表上单元格 I1 中第一个迷你图组的第一个迷你图的源数据的范围。*/
function test() {
    let sparkline = ActiveSheet.Range("I1").SparklineGroups.Item(1).Item(1)
    sparkline.SourceData = "A1:C1"
}
```
