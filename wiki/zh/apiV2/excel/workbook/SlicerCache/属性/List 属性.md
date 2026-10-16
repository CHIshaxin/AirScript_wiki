---
title: "List (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/List%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CRwtAWXm
---

# List (属性)

如果切片器缓存适用于表上的切片器，则为 **True**；否则为 **False**。 只读 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存是否适用于表上的切片器。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").List)
}
```

```javascript
/*本示例判断第一张工作簿第二个切片器缓存是否适用于表上的切片器，并通知用户。*/
function test() {
    let slCache = Workbooks.Item(1).SlicerCaches(2)
    if (slCache.List) {
        console.log("切片器缓存适用于表上的切片器")
    } else {
        console.log("切片器缓存不适用于表上的切片器")
    }
}
```
