---
title: "OLAP (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/OLAP%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DvIQvUG9
---

# OLAP (属性)

返回与指定的切片器缓存关联的切片器是否基于 OLAP 数据源。只读

## 说明

如果切片器是基于 OLAP 数据源，则为 **True**；否则为 **False**。

## 返回值

Boolean

## 示例

```javascript
/*本示例显示活动工作簿第二个切片器缓存关联的切片器是否基于 OLAP 数据源。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(2).OLAP)
}
```

```javascript
/*本示例判断如果第一张工作簿中名为“切片器_id”的切片器缓存关联的切片器未基于 OLAP 数据源，则删除该切片器缓存。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches("切片器_id")
    if (slcCache.OLAP == false) {
        slcCache.Delete()
    }
}
```
