---
title: "CacheIndex (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/CacheIndex%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CH065W4v
---

# CacheIndex (属性)

返回或设置数据透视表缓存的索引号。**Long** 类型，可读写。

## 说明

如果设置 **CacheIndex**属性，以便于数据透视表使用第二张数据透视表的高速缓存，则第一个报表的字段必须是第二个报表中字段的有效子集。

## 示例

```javascript
/*本示例将数据透视表“Pivot1”的缓存设置成数据透视表“Pivot2”的缓存。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").CacheIndex = Worksheets.Item(1).PivotTables("Pivot2").CacheIndex
}
```

```javascript
/*本示例显示数据透视表 Pivot1 缓存的索引号*/
function test() {
    console.log(Application.Worksheets.Item(1).PivotTables("Pivot1").CacheIndex)
}
```
