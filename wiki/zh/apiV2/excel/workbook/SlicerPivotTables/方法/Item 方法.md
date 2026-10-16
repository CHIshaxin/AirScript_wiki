---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerPivotTables/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Df_C76Zq
---

# Item (方法)

从集合中返回一个 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*本示例显示活动工作簿第一个切片器缓存关联的第一个数据透视表的名称。*/
function test() {
    let pvtTables = ActiveWorkbook.SlicerCaches(1).PivotTables
    console.log(pvtTables.Item(1).Name)
}
```

```javascript
/*本示例显示活动工作表上数据透视表的第一个切片器的缓存关联的第二个数据透视表的数据源。*/
function test() {
    let sliCache = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache
    console.log(sliCache.PivotTables.Item(2).SourceData)
}
```
