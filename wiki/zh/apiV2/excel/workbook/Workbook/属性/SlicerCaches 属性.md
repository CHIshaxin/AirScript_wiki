---
title: "SlicerCaches (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/SlicerCaches%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 9OuuKo-l
---

# SlicerCaches (属性)

表示与指定工作簿关联的切片器缓存的集合。

## 说明

使用 **[SlicerCaches](./../../SlicerCaches/SlicerCaches 对象.html)** 集合的 **Item** 属性可返回与指定的 Workbook 对象关联的 SlicerCache 对象。 可以使用 Index 属性的值，或者指定对象的 Name 属性来检索 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象。

## 返回值

SlicerCaches

## 示例

```javascript
/*本示例判断当活动工作簿关联的切片器缓存总数小于1，增加一个切片器缓存，否则通知用户第一个切片器缓存的名称。*/
function test() {
    if (ActiveWorkbook.SlicerCaches.Count < 1) {
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("数据透视表1"), "数量")
    } else {
        console.log(ActiveWorkbook.SlicerCaches.Item(1).Name)
    }
}
```

```javascript
/*本示例演示添加一个切片器缓存，并告知用户活动工作簿关联的切片器缓存集合的总数和第一个切片器缓存的名称。*/
function test() {
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("数据透视表1"), "日期")
    console.log(`活动工作簿关联的切片器缓存集合的总数为：${ActiveWorkbook.SlicerCaches.Count}\n活动工作簿中第一个切片器缓存的名称为：${ActiveWorkbook.SlicerCaches.Item(1).Name}`)
}
```
