---
title: "ListNames 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ListNames%20%E6%96%B9%E6%B3%95.html
---
# ListNames (方法)

从指定区域的第一个单元格位置开始，将所有未隐藏的名称的列表粘贴到工作表上。

## 说明

可以使用 **Names** 属性返回工作表中所有名称的集合。

## 返回值

Variant

## 示例

```javascript
/*本示例从 Sheet1 中单元格 A1 的位置开始，粘贴已定义名称的列表。本示例将粘贴 Sheet1 上定义的工作簿级名称和工作表级名称。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").ListNames()
}
```

```javascript
/*本示例将 B2 单元格名称设置为 testB，C1:C3 单元格名称设置为 testC，将工作表中所有未隐藏的名称以及相对应的地址显示到从 E1 单元格开始的区域*/
function test() {
    Range("B2").Name = "testB"
    Range("C1:C3").Name = "testC"
    Range("E1").ListNames()
}
```
