---
title: "SetIcon 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E6%96%B9%E6%B3%95/SetIcon%20%E6%96%B9%E6%B3%95.html
---
# SetIcon (方法)

为 **[SortField](./../../SortField/SortField 对象.html)** 对象设置图标。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Icon | Icon | 必填 | 要设置的图标。 |

## 示例

```javascript
/*本示例在活动工作表新建排序字段，并设置该字段按图标对数据进行排序。*/
function test() {
    let sortfield = Application.ActiveSheet.Sort.SortFields.Add(Range("C1:C5"), xlSortOnIcon)
    sortfield.SetIcon(ActiveWorkbook.IconSets(xl3Arrows).Item(1))
}
```

```javascript
/*本示例判断如果活动工作表排序字段按图标对数据进行排序，则设置图标类型为三标志。*/
function test() {
    let sortfield = Application.ActiveSheet.Sort.SortFields.Add(Range("A2:A5"), xlSortOnIcon)
    if (sortfield.SortON == xlSortOnIcon) {
        sortfield.SetIcon(ActiveWorkbook.IconSets(xl3Signs).Item(2))
    }
}
```
