---
title: "SortOn (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E5%B1%9E%E6%80%A7/SortOn%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cix7noRU
---

# SortOn (属性)

返回或设置要排序的单元格属性。可读/写 **[XlSortOn](./../../枚举/XlSortOn 枚举.html)** 类型。

## 说明

| XlSortOn 可以是下列 **[XlSortOn](./../../枚举/XlSortOn 枚举.html)** 常量之一： |
| --- |
| **SortOnCellColor** |
| **SortOnFontColor** |
| **SortOnIcon** |
| **SortOnValues** |

## 示例

```javascript
/*本示例在第一张工作表新建排序字段，并显示该排序字段是否按字体颜色对数据进行排序。*/
function test() {
    let sortfield = Application.Sheets.Item(1).Sort.SortFields.Add(Range("B1:B10"), xlSortOnFontColor)
    console.log(sortfield.SortOn == xlSortOnFontColor)
}
```

```javascript
/*本示例在活动工作表新建排序字段，并设置该排序字段根据单元格颜色进行排序。*/
function test() {
    let sortfield = Application.Sheets.Item(1).Sort.SortFields.Add(Range("B1:B10"))
    sortfield.SortON = xlSortOnCellColor
}
```
