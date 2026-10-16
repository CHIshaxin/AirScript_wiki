---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicers/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CXiDGoMO
---

# Item (方法)

从集合中返回一个 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | **[Slicer](./../../Slicer/Slicer 对象.html)** 对象的名称或索引号。 |

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_id”的切片器缓存中第一个切片器的高度。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_id").Slicers.Item(1).Height)
}
```

```javascript
/*本示例将活动工作表上数据透视表的第一个切片器内的列数设置为 2。*/
function test() {
    let slicer = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1)
    slicer.NumberOfColumns = 2
}
```
