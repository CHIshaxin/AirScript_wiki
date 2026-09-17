---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCaches/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

SlicerCache

## 示例

```javascript
/*本示例清除第一张工作簿中名为“切片器_id”切片器缓存的筛选器。*/
function test() {
    Workbooks.Item(1).SlicerCaches.Item("切片器_id").ClearManualFilter()
}
```

```javascript
/*本示例删除活动工作簿上名为“切片器_name”的切片器缓存，以及与该切片器缓存关联的“name”切片器。*/
function test() {
    ActiveWorkbook.SlicerCaches.Item("切片器_name").Delete()
}
```
