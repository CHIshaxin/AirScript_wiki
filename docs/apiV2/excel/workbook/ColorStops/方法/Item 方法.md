---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStops/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从指定集合中返回单个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

一个包含在集合中的 **[ColorStops](./../../ColorStops/ColorStops 对象.html)** 对象。

## 示例

```javascript
/*本示例删除活动单元格第三个ColorStop对象。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Item(3)
    colorStop.Delete()
}
```

```javascript
/*本示例将当前选定区域的第一个ColorStop对象的TintAndShade属性设置为-1。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Item(1)
    colorStop.TintAndShade = -1
}
```
