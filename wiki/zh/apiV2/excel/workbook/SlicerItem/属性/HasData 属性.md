---
title: "HasData (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/HasData%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CNOrTZkY
---

# HasData (属性)

返回切片器项是否包含与当前手动筛选状态相符的数据。只读。

## 说明

如果切片器项中包含与当前手动筛选状态相符的数据，则为 **True**；否则为 **False**。

此属性只在启用了交叉筛选的情况下才受支持，否则将生成运行时错误。若要确定是否为与数据透视表关联的切片器启用了交叉筛选，请检查指定的 **[SlicerItem](./../../SlicerItem/SlicerItem 对象.html)** 对象的父 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象的 **CrossFilterType** 属性值。若要确定是否为与 OLAP 数据源关联的切片器启用了交叉筛选，请检查对应于被筛选的 OLAP 层次结构的 **SlicerCacheLevel** 对象的 **CrossFilterType** 属性值。

## 返回值

Boolean

## 示例

```javascript
/*本示例显示活动工作簿第二个切片器缓存中第二个切片器项是否包含与当前手动筛选状态相符的数据。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(2).SlicerItems(2).HasData)
}
```

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存中名为“张五”的切片器项是否包含与当前手动筛选状态相符的数据。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").SlicerItems("张五").HasData)
}
```
